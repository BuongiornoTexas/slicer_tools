#!/usr/bin/env python
# cspell:ignore Bambu Fiberon Prusa
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
from enum import Enum, StrEnum
from typing import cast, NamedTuple, Generator
from zipfile import ZipFile
from zipfile import Path as zPath
from pathlib import Path

from openpyxl import Workbook  # type: ignore[import-untyped]

from supporting import AllNodeSettings, NodeMetadata, PresetType, PresetGroup
from supporting import SettingsDict, SettingValue
from supporting import FROM, INHERITS, NAME, DEFAULT_ENCODING
from slicer_presets import ProjectPresets

# Navigation and file name components
THREE_MF_FOLDER = "./3MF"
THREE_MF_EXT = ".3mf"
CONFIG_EXT = ".config"
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
INHERITS_GROUP = "inherits_group"  # Because why do one thing one way ...
VERSION = "version"


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
