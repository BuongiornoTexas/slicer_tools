#!/usr/bin/env python
# cspell:ignore Bambu Fiberon Bambulab Prusa
# TODO: Longer term: break this module into smaller sections.
# TODO: Figure out versioning in config files. This may be a problem for creating custom
# presets for import.
# TODO: Check todos in strings.
# TODO: Implement command line functions.
# TODO: Add special types to formatting (name, inherits, etc.) - also make version
# xx.xx.xx.xx.
# TODO: INCLUDE HUGE WARNING ABOUT INHERITING FROM USER DEFINED BASE TYPES (I
# create these as system types - they are very much not - see _user_base_fix).
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
    # preset is specified as the "inherits_group" value for the project overrride (i.e.
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


# TODO: #####CHECKED UP TO HERE#####
