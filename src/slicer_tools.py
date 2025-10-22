#!/usr/bin/env python
# cspell:ignore Bambu Fiberon Bambulab Prusa
# TODO: Longer term: break this module into smaller sections.
# TODO: Figure out versioning in config files. This may be a problem for creating custom
# presets for import.
# TODO: Implement command line functions.
# TODO: Add special types to formatting (name, inherits, etc.) - also make version
# xx.xx.xx.xx.
# TODO: INCLUDE HUGE WARNING ABOUT INHERITING FROM USER DEFINED BASE TYPES (I
# create these as system types - they are very much not - see _user_base_fix).
# TODO: Raise an issue about duplicate Fiberon presets.
"""Quick and dirty  tools for reading and comparing Bambu Studio slicer presets.

Summary: This package includes classes for loading system and user presets, and
presets/overrides specified in model .3mf files. The main feature of the package is
comparing presets in one or more .3mf files (the default mode summarises differences
relative to system presets). Differences are provided in .xlsx format. In addition, all
settings in a single preset can be dumped as a json file (from walking the inheritance
tree).

Notes:
    - In this package, I focus on filament, machine and process as the core preset
    types.
    - Preset types may be system, user and project (these three follow the BambuStudio
    naming convention) and **override**, which is a preset generated from project file
    overrides of one of three other groups. Override is not BambuStudio terminology,
    but provides a convenient naming convention for handling project file oddities.

Cautions:
    - This implementation has only been tested on  Windows. But if the Bambu Studio file
    structure for macos mimics Windows, it should work OK on there too (or at least
    can be made to work with minimal updates).
    - I have only implemented an english version, so this package may break when using
    other languages.
    - Right now the package is built around Bambu Studio files. However, I expect that
    it should be easy to extend to Orca Slicer files, and possibly also Prusa Slicer.
"""
import json
import string
import random
from copy import deepcopy
from enum import Enum, StrEnum
from sys import platform
from typing import cast, NamedTuple, Generator
from zipfile import ZipFile
from zipfile import Path as zPath
from pathlib import Path

from openpyxl import Workbook  # type: ignore[import-untyped]

# Navigation and file name components
THREE_MF_FOLDER = "./3MF"
THREE_MF_EXT = ".3mf"
CONFIG_EXT = ".config"
DEFAULT_ENCODING = "utf-8"
METADATA = "Metadata/"

# Constants for manipulating .configs
# Can build all json file names from these two constants and the
# PresetType enum below.
SETTINGS = "_settings"
PROJECT_CONFIG_NAME = "project" + SETTINGS
EXCLUDE_CONFIGS = ["model_settings", "slice_info"]

# more json key elements
# annoyingly, in config files, settings id uses a different terminology to
# the file naming - specifically print in the json is equivalent to process and
# printer is equivalent to machine. Yay!
PRINT = "print"
PRINTER = "printer"
DIFFS_TO_SYSTEM = "different_settings_to_system"
ID = "_id"
INHERITS = "inherits"
INHERITS_GROUP = "inherits_group"  # Because why do one thing one way ...
FROM = "from"
NAME = "name"
VERSION = "version"

# create some type aliases
type SettingValue = str | list[str]
type SettingsDict = dict[str, SettingValue]


class CellFormat(StrEnum):
    """Builtin Excel cell formats."""

    NORMAL = "Normal"
    GOOD = "Good"
    BAD = "Bad"
    INPUT = "Input"
    NOTE = "Note"
    NEUTRAL = "Neutral"
    HEADING4 = "Headline 4"


class DiffType(Enum):
    """Difference types."""

    # No difference in values.
    NO_DIFF = 0
    # Reference value (the value differenced to!).
    REFERENCE = 1
    # Normal difference to reference value.
    DIFFERENCE = 2
    # Override system difference.
    OVERRIDE = 3
    # Unset reference value.
    UNSET = 4
    # Too complex to diff (for now at least).
    COMPLEX = 5


class CellInfo(NamedTuple):
    """Summary data for writing to an Excel cell."""

    row: int
    column: int
    value: str
    format: DiffType | CellFormat


class PresetType(StrEnum):
    """Preset types used by Bambu Studio."""

    FILAMENT = "filament"
    MACHINE = "machine"
    PROCESS = "process"


class PresetGroup(StrEnum):
    """Preset groups used by Bambu Studio.

    System refers to presets from the system folder.
    User refers to presets from the user folder.
    Project refers to a filament, process or machine preset in a 3mf file.
    Override refers to the project settings from a 3mf file that override any of
    of the above (this is strictly also project preset, but as it is handled somewhat
    differently to the other presets by BBS, I've created a special category).
    """

    OVERRIDE = "override"
    PROJECT = "project"
    USER = "user"
    SYSTEM = "system"


class PresetPath(NamedTuple):
    """Provide a unique path/key for presets in ProjectPresets.

    This tuple eliminates the requirement sparse nested dictionaries or minimal classes
    emulating nested dicts (at the cost of using tuple hashes for dict keys).
    """

    # Note that this path should be guaranteed to be unique within a project (.3mf)
    # file, but is NOT guaranteed unique across files. See DiffValuePath to deal with
    # this problem.
    type: PresetType
    preset_name: str


class NodeMetadata(NamedTuple):
    """Summary metadata relating to Preset Nodes.

    Used for initialising PresetNodes and in generating complete settings lists from
    tree walk.
    """

    # Unique name for preset in project file.
    name: str
    # Source filename for the preset. Source file may contain multiple presets.
    filename: str
    group: PresetGroup
    preset_type: PresetType

    # The override_inherits field deals with an irritating quirk of project
    # overrides - they are defined as differences to a system preset. This system
    # preset is specified as the "inherits_group" value for the project override (i.e.
    # the parent for the differences is relative to the system preset named in
    # "inherits_group").
    #
    # However, in the Bambu Studio interface, orange highlighted values show differences
    # relative to the project specified immediate parent of the override, which may not
    # be the same as the system preset above! The immediate parent is named in:
    # "<printer|print|filament>_settings_id". (i.e. the project settings includes two
    # inheritance paths for overrides!)
    #
    # To capture this information in this package, we need to know both the system
    # preset and the immediate parent preset.
    #
    # I could have handled this by introducing a dummy key/value pair into the
    # the settings dictionary, but this would probably cause bigger problems with
    # end users trying to create a preset containing this dummy parameter.
    #
    # So instead we do a special here to handle the unicorn sparkly tears case
    # of project settings.
    #
    # override_inherits is the name of the immediate parent preset for the override. It
    # may be different from the "inherits" value used to generate
    # "differences_to_system" values.
    #
    # Defaults to empty string and should only be set for group = Preset.OVERRIDE.
    override_inherits: str = ""


class DiffValue(NamedTuple):
    """Container for difference value and difference type information."""

    value: str
    type: DiffType


class DiffValuePath(NamedTuple):
    """Provide a unique path/key for difference sets.

    This tuple eliminates the requirement for sparse nested dictionaries or minimal
    classes emulating nested dicts, at the cost of using tuples as dict keys.
    """

    # This is pretty similar to PresetPath. However, a diff set may contain multiple
    # preset types and multiple .3mf files, so we need to add  the filename to ensure
    # uniqueness - as this info is already available in node metadata, we're just using
    # that. In addition, the value (row) name will also be sparse. As adding row name
    # to the key makes iteration over differences easier, I've done that as well.
    row_name: str
    metadata: NodeMetadata


class DiffMatrix:
    """Difference data in a sparse matrix form."""

    # This started as a dataclass, but as all members are now private, converted to
    # full fledged class.
    _rows: dict[str, int]
    _cols: dict[NodeMetadata, int]
    _values: dict[DiffValuePath, DiffValue]
    # Flag indicating indices need recalculating. Set to true every time a new
    # value is added.
    _reset_required: bool
    # Group, filename, preset name.
    _header_count: int = 3

    def __init__(self) -> None:
        """Create instance variables."""
        self._rows = {}
        self._cols = {}
        self._values = {}
        self._reset_required = True

    def column_exits(self, metadata: NodeMetadata) -> bool:
        """Test if column is defined."""
        if metadata in self._cols:
            return True
        return False

    def add_value(
        self, row_name: str, column_id: NodeMetadata, value: str, value_type: DiffType
    ) -> None:
        """Add or overwrite a value to the diff matrix.

        Automatically triggers an instance reset and will make any active generator
        invalid (values/indices have been modified in ways that are unpredictable for
        the generator).
        """
        self._reset_required = True
        self._rows[row_name] = -1
        self._cols[column_id] = -1
        self._values[DiffValuePath(row_name, column_id)] = DiffValue(value, value_type)

    def _reset_lookups(self) -> None:
        """Prepare row and column lookup indices for writing tables.

        All indices are 0 based.
        """
        key: str | NodeMetadata

        keys = self._rows.keys()
        for i, key in enumerate(sorted(keys), start=0):
            self._rows[key] = i

        # Column sort is not worth doing, as we want a rough sort along the lines:
        #   Reference, Override, Project and User.
        # While it would be possible to write a compare function, it's nearly as
        # simple to just run over the column list for each group.
        i = 0
        for group in [
            PresetGroup.SYSTEM,
            PresetGroup.OVERRIDE,
            PresetGroup.PROJECT,
            PresetGroup.USER,
        ]:
            for key, _ in self._cols.items():
                if key.group == group:
                    self._cols[key] = i
                    i += 1

    def table_cells(self) -> Generator[CellInfo]:
        """Generate table and associated row names and column headers.

        Return values are row offset (0 based), column number, value, value type or
        format.
        """
        if self._reset_required:
            self._reset_lookups()

        # Header names. All 0 based.
        yield CellInfo(0, 0, "Group", CellFormat.NORMAL)
        yield CellInfo(1, 0, "Filename", CellFormat.NORMAL)
        yield CellInfo(2, 0, "Preset", CellFormat.NORMAL)

        for row_name, row_offset in self._rows.items():
            yield CellInfo(
                row_offset + self._header_count, 0, row_name, CellFormat.NORMAL
            )

        for metadata, col_offset in self._cols.items():
            yield CellInfo(0, col_offset+1, metadata.group.value, CellFormat.NORMAL)
            yield CellInfo(1, col_offset+1, metadata.filename, CellFormat.NORMAL)
            yield CellInfo(2, col_offset+1, metadata.name, CellFormat.NORMAL)

        for key, value in self._values.items():
            row = self._rows[key.row_name] + self._header_count
            col = self._cols[key.metadata] + 1
            yield CellInfo(row, col, value.value, value.type)

    def row_count(self) -> int:
        """Row count in the table include header lines."""
        return len(self._rows) + self._header_count


class AllNodeSettings(NamedTuple):
    """Container for all of the settings for a preset node (from tree roll up)."""

    # metadata for the source/target node
    metadata: NodeMetadata
    # metadata for the reference node if defined.
    ref_metadata: NodeMetadata | None
    # Settings generated by walking the inheritance tree from the source/target node
    # up to but not including the reference node if defined, or all settings if
    # reference node is not defined).
    # In effect, these are the source node settings that are different to the reference
    # node settings.
    # (If reference settings exist, these settings either override or augment the
    # reference settings.)
    source_subtree: SettingsDict
    # Settings generated by walking the inheritance tree from the reference node to
    # the root of the inheritance tree.
    reference_subtree: SettingsDict


class PresetNode:
    """Container for preset or override definition.

    A preset DataNode contains all of the information required for preset/override
    definition.

    The primary data managed by this class is "the collection of key/value pairs for any
    settings" that differ from "key/value pairs the preset inherits from its immediate
    parent". That is, a Preset effectively represents the difference in settings between
    the source preset and the parent of the source preset.

    A Preset does not address any differences further up the inheritance tree - these
    are generated by walking the inheritance tree (addressed in PresetNodes container
    class.

    A PresetNode contains settings for single preset type, and may be sourced from a
    json preset file, a preset config file in a 3mf archive, or a preset override from a
    project config file in a 3mf archive.

    Intended as a read-only source for settings key/value pairs and other identity
    parameters.
    - filename is the name of the source file, which may be a normal file or a file
    in a 3mf archive.
    - name is the name of the preset. Note that ProjectPresets (a container class for
    Presets) requires that each name in a 3mf project is unique. See ThreeMFPresets for
    a sample implementation to manage this.
    - path is the path to the source preset file if the source is a normal file, None
    if the source is a 3mf archive member.
    """

    metadata: NodeMetadata
    path: Path | None = None
    # Settings is a private variable to support lazy loading of system and user
    # preset settings.
    _settings: SettingsDict | None

    # This is a REAL hack. BambuStudio more or less treats configs in
    # base directories in user folders as more-or-less system configs.
    # We use this flag to modify these on the fly to more-or-less follow
    # the way BS behaves.
    _user_base_fix: bool = False

    def __init__(
        self,
        metadata: NodeMetadata,
        settings: SettingsDict | None = None,
        path: Path | None = None,
        user_base_fix: bool = False,
    ) -> None:
        """Create preset instance."""
        # PresetNode is designed around lazy loading system and user settings.
        # If the default value/None is provided for settings, lazy loading will
        # be assumed.
        # self._settings is initialised here, and lazy loading occurs in the data
        # getter.
        self._settings = settings
        self.path = path
        if user_base_fix:
            self._user_base_fix = user_base_fix
            # Faux system group in the base folder. Second hack below.
            self.metadata = metadata._replace(group=PresetGroup.SYSTEM)
        else:
            self.metadata = metadata

        if self._settings is None and self.path is None:
            # Shouldn't happen, but in case.
            raise ValueError(
                "Settings and path arguments cannot both be None for a PresetNode()"
            )

    @property
    def settings(self) -> SettingsDict:
        """Return the node settings dictionary."""
        # see __init__ for an explanation of _settings and the settings getter.
        if self._settings is None:
            if self.path is None:
                # I think this should be caught during initialisation. But belt and
                # braces.
                raise KeyError(
                    "Preset node contains no settings data and no path to preset file."
                )

            # Lazy load needed for settings.
            with open(self.path, encoding=DEFAULT_ENCODING) as fp:
                self._settings = json.load(fp)
                if self._user_base_fix:
                    # This is our second hack to make user/**/base/*.json look
                    # like system presets.
                    self._settings[FROM] = PresetGroup.SYSTEM

        # Relying on the caller to not modify settings.
        return self._settings


class ProjectPresets:
    """Container for a group of related Preset nodes built around a BS .3mf model.

    ProjectPresets provides access to a related group of presets and overrides, both at
    the level of nodes in the inheritance tree, and as roll ups of the inheritance
    tree for a node. As the main preset grouping I'm working with is a BambuStudio
    .3mf project file, so this class is based around a .3mf structure.

    The container manages two groups of Preset nodes:
    - Shared preset nodes available to all instances of ProjectSettings. These are:
        - BambuStudio system presets.
        - BambuStudio user presets for one user, which defaults to the last user to
        login to BambuStudio.
        - Note the shared presets will be a superset of the system/user presets used by
        any given project file, but it's convenient to grab them all at once.
    - A set of per instance project preset nodes, corresponding to:
        - Project (3mf file unit) settings for filament, machine and process presets.
        - Project overrides of any of the above presets.
    The shared settings load on the first instantiation of ProjectSettings, and can only
    be overridden on this first call.
    """

    # Both preset node dictionaries are keyed by PresetPath named tuples.
    # Should be enough for relatively efficient returns, but DOES require creation of
    # unique names for project overrides of presets.
    _project_nodes: dict[PresetPath, PresetNode]
    # Class variable, share among all instances.
    _shared_nodes: dict[PresetPath, PresetNode] = {}

    def __init__(self, appdata_path: Path | None = None, bbl_user_id: str = "") -> None:
        r"""Create node containers and load system and user preset nodes.

        appdata_path and bbl_user_id will only be used in the first instance call, and
        ignored thereafter.

        appdata_path points to the location for the Bambu Studio system snd user
        preset nodes, and is the source for _shared_settings.
        If not specified, this defaults to
            Windows: c:\users\<username>\appdata\Roaming\BambuStudio
            macos: /Users/user.name/Library/Application Support/BambuStudio
        I don't have a mac, so haven't tested the macos version.

        If specified, the folder structure must follow the form:
            appdata_path
                BambuStudio.conf (only required if bbl_user_id is not specified)
                system
                    BBL.json (must follow the same format as the default version)
                    BBL
                        filament
                        machine
                        process
                user
                    <bbl_user_id>
                        filament
                            base
                        machine
                            base
                        process
                            base
        bbl_user_id is the user id of the folder containing user presets. If it is not
        specified, a default value will be extracted from BambuStudio.conf (this will
        typically be the id of the last user to login to BS). Initialisation will fail
        dramatically if user_id is not specified and BambuStudio.conf does not exist.
        """
        if len(ProjectPresets._shared_nodes) == 0:
            # First instance of this class, shared settings don't exist.
            self._load_shared_nodes(appdata_path, bbl_user_id)

        # Create instance project dict.
        self._project_nodes = {}

    @classmethod
    def _load_shared_nodes(
        cls, appdata_path: Path | None = None, bbl_user_id: str = ""
    ) -> None:
        """Load shared nodes dictionary."""
        # Given this is the only place these paths should appear, happy enough hard
        # coding here.
        studio_path = appdata_path
        if studio_path is None:
            if platform == "win32":
                studio_path = Path.home() / "Appdata/Roaming"
            elif platform == "darwin":
                studio_path = Path.home() / "Library/Application Support"
            else:
                raise RuntimeError(
                    f"BambuScan doesn't know how to find BambuStudio"
                    f" folder on {platform}"
                )
            studio_path = studio_path / "BambuStudio"

        user_id = bbl_user_id
        if not user_id:
            try:
                with open(
                    studio_path / "BambuStudio.conf", encoding=DEFAULT_ENCODING
                ) as fp:
                    # strip out checksum comment. Hells bells.
                    raw_json = "".join(line for line in fp if not line.startswith("#"))
            except FileNotFoundError:
                print("Missing BambuStudio.conf file.")
                raise

            json_data = json.loads(raw_json)
            try:
                user_id = json_data["app"]["preset_folder"]
            except KeyError:
                print("Missing 'preset folder' key/value in BambuStudio.conf")
                raise

        # Populate system presets.
        cls._system_nodes_from_bbl_json(studio_path / "system")

        # Walk paths. Should just verify system, loads user.
        # Duplicates will get a printed warning, but no error.
        cls._nodes_from_path_walk(
            root=studio_path / "system/BBL", group=PresetGroup.SYSTEM
        )
        cls._nodes_from_path_walk(
            root=studio_path / "user" / user_id, group=PresetGroup.USER
        )

        # And just in case we didn't find anything, add a dummy entry so
        # we don't go through this on each load.
        # Likely the caller will find out the hard way.
        if len(cls._shared_nodes) == 0:
            cls._shared_nodes[PresetPath(PresetType.FILAMENT, "dummy")] = PresetNode(
                NodeMetadata(
                    name="Dummy",
                    filename="Dummy.json",
                    group=PresetGroup.SYSTEM,
                    preset_type=PresetType.PROCESS,
                ),
                settings={"name": "No standard settings found"},
            )

    @classmethod
    def _system_nodes_from_bbl_json(cls, system_path: Path) -> None:
        """Load system preset file locations from BBL.json.

        BBL.json is assumed definitive. I know there are duplicates. E.g. Fiberon.
        We will warn about these and ignore the duplicates.
        """
        # Some hard coding of strings that should only occur here.
        with open(system_path / "BBL.json", encoding=DEFAULT_ENCODING) as fp:
            data = json.load(fp)

        for preset_type in PresetType:
            # ignoring machine models data.
            for item in data[preset_type + "_list"]:
                name = item[NAME]
                path = system_path / "BBL" / item["sub_path"]

                preset_key = PresetPath(preset_type, name)
                if preset_key in cls._shared_nodes:
                    raise ValueError(
                        # Warn, but don't do anything about it.
                        f"Highly unexpected duplicate definition in BBL.json:"
                        f"   '{preset_key}'."
                    )

                # Prep for lazy load.
                cls._shared_nodes[preset_key] = PresetNode(
                    NodeMetadata(
                        name=name,
                        filename=path.name,
                        group=PresetGroup.SYSTEM,
                        preset_type=preset_type,
                    ),
                    path=path,
                )

    @classmethod
    def _nodes_from_path_walk(cls, root: Path, group: PresetGroup) -> None:
        """Load preset nodes from path walk. Really for user paths."""
        for preset_type in PresetType:
            base_path = root / preset_type
            for json_path in base_path.glob("**/*.json"):
                preset_key = PresetPath(
                    type=preset_type,
                    preset_name=json_path.stem,
                )
                if preset_key in cls._shared_nodes:
                    if json_path != cls._shared_nodes[preset_key].path:
                        # Warn but don't do anything about it.
                        print(f"Warning duplicate definition for '{preset_key}`.")

                else:
                    # Tell PresetNode to deal with user bases that are actually
                    # system types. (Handled by Preset Node as multiple fixes required.)
                    base_fix = False
                    if group == PresetGroup.USER and json_path.parent.name == "base":
                        base_fix = True

                    cls._shared_nodes[preset_key] = PresetNode(
                        NodeMetadata(
                            name=json_path.stem,
                            filename=json_path.name,
                            group=group,
                            preset_type=preset_type,
                        ),
                        path=json_path,
                        user_base_fix=base_fix,
                    )

    def add_project_node(
        self,
        metadata: NodeMetadata,
        settings: SettingsDict,
    ) -> None:
        """Add project preset node based on the data dict."""
        key = PresetPath(
            type=metadata.preset_type,
            preset_name=metadata.name,
        )

        # Quick and dirty checks on uniqueness.
        if key in self._shared_nodes:
            raise KeyError(
                f"Attempt to add project preset {key} with the same name as a system"
                f" preset."
            )
        if key in self._project_nodes:
            raise KeyError(
                f"Attempt to add project preset {key} when preset already exists."
            )
        # Arguably should check name is unique, but leaving that up to BS for now.
        # preset_type = settings[]
        self._project_nodes[key] = PresetNode(
            metadata=metadata,
            settings=settings,
        )

    def _node(self, preset_type: PresetType, name: str) -> PresetNode:
        """Return the specified preset node."""
        key = PresetPath(type=preset_type, preset_name=name)
        if key in self._shared_nodes:
            return self._shared_nodes[key]

        if key in self._project_nodes:
            return self._project_nodes[key]

        raise KeyError(f"Undefined preset '{key}'.")

    def node_settings(self, preset_type: PresetType, name: str) -> SettingsDict:
        """Return a **deep copy** of the node settings named preset node.

        The settings returned by this call are the differences between the settings
        for the named node and the settings of the parent node (i.e. the preset data
        that would appear in the json file for the preset).
        """
        # Keep the source data read only.
        return deepcopy(self._node(preset_type, name=name).settings)

    def all_node_settings(
        self,
        preset_type: PresetType,
        node_name: str,
        ref_node: str = "",
        ref_group: PresetGroup | None = None,
    ) -> AllNodeSettings:
        """Return all settings for node_name of preset_type.

        Settings may be provided as either
        - a single SettingsDict and associated NodeMetadata for node_name (default), or
        - as partitioned settings consisting of a SettingsDict generated by walking the
        inheritance from node_name up to but not including the reference node, and a
        reference settings dictionary from walking the inheritance tree from the
        reference node to the root of the tree. (i.e. reference settings and
        node_name settings that are different to the reference settings).

        The reference node is determined from either the reference node name ref_node
        or the youngest inherited ancestor of node_name that has a PresetGroup of
        ref_group. At least one of ref_value and ref_group must be empty ("").

        If called with default settings, the return structure is:
        - ret_val.metadata contains metadata for node_name.
        - ret_val.settings contains all settings for node_name.
        - ret_val.ref_metadata = None.
        - ret_val.reference = {}

        If called with with a reference group/node specified, the return structure is:
        - ret_val.metadata contains metadata for node_name.
        - ret_val.settings contains all settings different to the reference settings.
        - ret_val.ref_metadata contains the metadata for the reference node.
        - ret_val.reference contains all settings for the reference node.
        """
        if ref_group and ref_node:
            raise ValueError(
                f"Cannot specify both reference group '{ref_group}'"
                f" and reference node '{ref_node}'"
            )

        # Unfortunately because of the way python dictionary unions work (keep
        # rightmost value), we need to walk the tree first and then build the
        # return dictionaries with union assignment |= from the root back to the
        # calling node. I've opted to keep this in a a single method, as I think the
        # logic is cleaner than via recursion.
        # First up create node lists for the settings/different to reference settings
        # and the reference settings.
        settings_nodes: list[SettingsDict] = []
        reference_nodes: list[SettingsDict] = []
        # Collect settings until we find the first reference node and then
        # switch.
        target = settings_nodes

        # Current node for processing.
        this_node: PresetNode | None = self._node(preset_type, node_name)
        # metadata for this node.
        metadata: NodeMetadata = cast(PresetNode, this_node).metadata
        ref_metadata: NodeMetadata | None = None

        while this_node:
            # This next bit looks stupid, but I think it's correct.
            # The first thing we do is check if we need to switch to gathering reference
            # nodes. This may have the odd side effect that the original node may
            # be assigned directly to the reference settings if it is in the reference
            # group.
            if not ref_metadata and (
                (ref_group and this_node.metadata.group == ref_group)
                or (this_node.metadata.name == ref_node)
            ):
                # This is either the youngest ancestor in reference group, or this is
                # the reference node, so we switch to gathering reference data.
                ref_metadata = this_node.metadata
                target = reference_nodes

            # Warning - we are grabbing mutable node SettingsDicts here.
            # Caller should not alter these.
            target.append(this_node.settings)

            # Move on to parent node.
            # By rights inherits should ALWAYS be a str, but sometimes doesn't exist at
            # all.
            if INHERITS not in this_node.settings:
                inherits = ""
            else:
                inherits = cast(str, this_node.settings[INHERITS])

            if inherits:
                this_node = self._node(preset_type, inherits)
            else:
                this_node = None

        # Prepare empty settings dicts and
        # unroll the lists to populate the dicts, and finally sort them.
        working: SettingsDict = {}
        for node in reversed(settings_nodes):
            working |= node
        settings = dict(sorted(working.items()))
        working = {}
        for node in reversed(reference_nodes):
            working |= node
        reference = dict(sorted(working.items()))

        # I don't think this is possible, but no harm in a quick check.
        if len(reference) != 0 and not ref_metadata:
            raise ValueError(
                "Reference values defined, but reference metadata is not."
                "\nThis should not be possible."
            )

        return AllNodeSettings(
            metadata=metadata,
            ref_metadata=ref_metadata,
            source_subtree=settings,
            reference_subtree=reference,
        )

    def project_presets(
        self, reference_group: PresetGroup | None = None, reference_node: str = ""
    ) -> Generator[AllNodeSettings]:
        """Yield settings and metadata for project presets relative to reference group.

        The generator repeatedly calls all_node_settings to generate and return
        the setting for each project preset node in the ProjectPresets instance.

        If reference group or reference_node is specified, these settings are all
        relative to the reference node/group (see all_node_settings for more detail),
        which is also returned in AllNodeSettings.
        """
        for node in self._project_nodes.values():
            # Dom't need keys as we iterate over all elements in _project_nodes,
            # and key data is repeated in node.metadata.
            yield self.all_node_settings(
                node.metadata.preset_type,
                node_name=node.metadata.name,
                ref_group=reference_group,
                ref_node=reference_node,
            )


class ThreeMFPresets:
    """Container for all presets in a 3mf file.

    Currently skips per object and slice data. And will probably never implement these.
    """

    filename: str
    # Project settings extracted from project_settings.config.
    project_config: SettingsDict

    # Collection of the following presets:
    # - All project presets and overrides in the project file.
    # - All user presets that the project presets/overrider might inherit from.
    # - All system presets that the project presets/overrides might inherit from.
    # ThreeMFPresets creates the first group, the other two are loaded automatically.
    # After all project presets have been created, the project_presets member
    # can generate settings and differences as required.
    # The caveat is that calls to generate settings/differences must be deferred until
    # all Presets in the 3mf have been instantiated, as otherwise there is a risk that
    # the project_presets object will not hold the data required for a complete roll up.
    # (In effect this is creating a half way house to a class variable that is only
    # visible within each ThreeMFSettings instance and the Settings instances contained
    # by it. Yuck!)
    presets: ProjectPresets

    def __init__(
        self, project_file: Path, appdata_path: Path | None = None, user_id: str = ""
    ) -> None:
        """Create project metadata container.

        appdata_path and user_id override the default locations for find the system
        and user presets.
        """
        self.filename = project_file.name
        self.presets = ProjectPresets(appdata_path=appdata_path, bbl_user_id=user_id)
        with ZipFile(project_file, mode="r") as archive:
            for zip_path in zPath(archive, at=METADATA).iterdir():
                # Iterate over metadata folder to find presets (.config files).
                if (
                    zip_path.name.endswith(CONFIG_EXT)
                    and zip_path.stem not in EXCLUDE_CONFIGS
                ):
                    # Excluding xml format configs for now.
                    # We'll need the config data no matter what.
                    # By rights, this should be a json file at this point.
                    settings = json.loads(zip_path.read_text(encoding=DEFAULT_ENCODING))
                    if zip_path.stem == PROJECT_CONFIG_NAME:
                        # Project config. Will contain multiple overrides and
                        # requires special handling.
                        self.project_config = settings
                        self._process_project_settings()
                    else:
                        # Find out what type we are working with and create preset.
                        for prefix in PresetType:
                            if zip_path.stem.startswith(prefix + SETTINGS):
                                self.presets.add_project_node(
                                    NodeMetadata(
                                        name=settings[NAME],
                                        filename=project_file.name,
                                        preset_type=prefix,
                                        group=PresetGroup.PROJECT,
                                    ),
                                    settings=settings,
                                )
                                break

    def _process_project_settings(self) -> None:
        """Extract presets from project_settings.config and add to project presets."""
        # process the differences by type.
        # Guaranteed there is a better way to this, but patience is gone.
        group = PresetGroup.OVERRIDE
        filament_count = len(self.project_config[PresetType.FILAMENT + SETTINGS + ID])
        filament_idx = -1
        for diff_idx, diff in enumerate(self.project_config[DIFFS_TO_SYSTEM]):
            if diff_idx == 0:
                # Process.
                name = self.project_config[PRINT + SETTINGS + ID]
                value_idx = 0
                preset_type = PresetType.PROCESS

            elif diff_idx == filament_count + 1:
                # Machine.
                name = self.project_config[PRINTER + SETTINGS + ID]
                value_idx = 0
                preset_type = PresetType.MACHINE

            else:
                # Filament
                filament_idx += 1
                name = self.project_config[PresetType.FILAMENT + SETTINGS + ID][
                    filament_idx
                ]
                value_idx = filament_idx
                preset_type = PresetType.FILAMENT

            # Set up the special unicorn tears preset.
            # Create settings and add key/values missing from the override differences.
            # Unfortunately, the project override name is not unique, as it is also the
            # the name of the immediate parent that the override inherits from (see
            # PresetMetadata for more on this, and why we are setting up the fix we
            # do here).
            # The fix is to create the preset metadata with the parent preset inherit
            # specified and the preset name made unique by adding 4 random characters
            # to the end of the name.
            name = cast(str, name)
            metadata = NodeMetadata(
                name=name + "_" + "".join(random.sample(string.ascii_lowercase, 4)),
                filename=self.filename,
                preset_type=preset_type,
                group=group,
                # This is the immediate parent, which may be different from the
                # inherits_group value in project_settings. Yay!
                override_inherits=name,
            )

            # And now we pre-populate settings with values that project_settings does
            # not provide. As another project_settings wrinkle: if the project does not
            # override the system values at all, "inherits" is set to "".
            # We do some munging here to fix this as well. Yay.
            inherits = self.project_config[INHERITS_GROUP][diff_idx]
            if not inherits:
                inherits = name

            settings = {
                # Use unique name for override.
                NAME: metadata.name,
                INHERITS: inherits,
                FROM: self.project_config[FROM],
                VERSION: self.project_config[VERSION],
            }

            # Finally collect the actual differences.
            self._differences_to_settings(
                value_idx=value_idx,
                filament_count=filament_count,
                diff=diff,
                settings=settings,
                metadata=metadata,
            )

            self.presets.add_project_node(
                metadata=metadata,
                settings=settings,
            )

    def _differences_to_settings(
        self,
        value_idx: int,
        diff: str,
        settings: SettingsDict,
        metadata: NodeMetadata,
        filament_count: int = 0,
    ) -> None:
        """Create override settings from difference list.

        expect_count is the number of values expected in the value lists. Used for
        list length validations. List length validation tries to be smart, but may get
        things wrong. Check errors and warnings to resolve.
        """
        diff_keys = diff.split(";")
        if not diff_keys[0]:
            return

        # If I was confident about value counts, could do this with a comprehension.
        # But instead we'll do it slow with checks. And we need to, because of
        # inconsistency.
        for key in diff_keys:
            values = self.project_config[key]

            # Deal with filament first.
            if metadata.preset_type == PresetType.FILAMENT:
                if key == "filament_notes":
                    print(
                        f"Warning: key `filament_notes' appears in override"
                        f" {metadata.name}. Key ignored."
                        f"\n  (BambuStudio currently retains the note value for one"
                        f"  filament only in the config file and I have no idea which"
                        f"  one is important.)"
                    )
                    continue

                if not isinstance(values, list):
                    raise TypeError(
                        f"Expected list of values for setting '{key}'"
                        f"\n  of filament override {metadata.name}."
                        f"\n  Got {type(values)}. (Values = {values})."
                    )

                if len(values) != filament_count:
                    # Will raise a warning or error. Prep info string accordingly.
                    raise IndexError(
                        f"Unexpected override list size."
                        f"\n  List length is expected to match override/project"
                        f" filament count."
                        f"\n  Override `{metadata.name}` setting `{key}`:"
                        f"\n     Expected list length count {filament_count}, got"
                        f" {len(values)}."
                        f"\n    (Values = {values})"
                    )

                # If we're here, I think it's real.
                settings[key] = values[value_idx]

            else:
                # Now dealing with machine and process values. These could be anything,
                # but typically only expect either a string or single value list.
                # Warn about anything unexpected and grab the value anyway.
                if isinstance(values, str) or (
                    isinstance(values, list) and len(values) == 1
                ):
                    # What we expected.
                    settings[key] = values

                else:
                    print(
                        f"Warning: Unexpected type for machine/process override."
                        f" This is probably OK, but should be checked to make sure."
                        f"\n  In override `{metadata.name}` for setting `{key}`:"
                        f"\n  Expected either a single value list or a string."
                        f"\n  Got (Values = {values})"
                    )
                    # Not what we expected, but still probably fine.
                    settings[key] = values


# TODO: #####CHECKED UP TO HERE#####
