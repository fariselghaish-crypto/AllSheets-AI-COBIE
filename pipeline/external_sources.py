
import io
import pandas as pd

EMPTY_VALUES = {"", "TBC", "nan", "None", "NaN", "NONE"}


def load_external_database(uploaded_file):
    if uploaded_file is None:
        return {}

    try:
        xls = pd.ExcelFile(uploaded_file)
        data = {}
        for sheet in xls.sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name=sheet)
            if not df.empty:
                df.columns = [str(c).strip() for c in df.columns]
                data[str(sheet).strip()] = df
        return data
    except Exception:
        return {}


def _is_empty(value):
    return str(value).strip() in EMPTY_VALUES


def _find_match(row, ext_df):
    if ext_df is None or ext_df.empty:
        return None

    keys = [
        "Name",
        "TypeName",
        "AssetIdentifier",
        "ExternalIdentifier",
        "RowName",
    ]

    for key in keys:
        if key in row.index and key in ext_df.columns:
            val = str(row.get(key, "")).strip()
            if not val or val in EMPTY_VALUES:
                continue
            matches = ext_df[ext_df[key].astype(str).str.strip() == val]
            if not matches.empty:
                return matches.iloc[0]

    return None


def merge_external_sources(cobie_sheets, external_sheets):
    if not external_sheets:
        return cobie_sheets, ["No external database uploaded. Generated COBie data only."]

    logs = []
    merged = {}

    for sheet_name, df in cobie_sheets.items():
        df2 = df.copy()
        ext_df = external_sheets.get(sheet_name)

        if ext_df is None or ext_df.empty:
            merged[sheet_name] = df2
            logs.append(f"{sheet_name}: no external sheet found.")
            continue

        updates = 0

        for idx, row in df2.iterrows():
            match = _find_match(row, ext_df)
            if match is None:
                continue

            for col in df2.columns:
                if col in ext_df.columns:
                    current = df2.at[idx, col]
                    external = match.get(col)
                    if _is_empty(current) and not _is_empty(external):
                        df2.at[idx, col] = external
                        updates += 1

        merged[sheet_name] = df2
        logs.append(f"{sheet_name}: {updates} fields updated from external database.")

    return merged, logs


def external_source_template():
    sheets = {
        "Facility": ["Name", "ProjectName", "SiteName", "Category", "Description"],
        "Floor": ["Name", "Category", "Elevation", "Height", "Description"],
        "Space": ["Name", "FloorName", "Category", "GrossArea", "NetArea", "UsableHeight"],
        "Zone": ["Name", "Category", "SpaceNames", "Description"],
        "Type": ["Name", "Category", "Description", "AssetType", "Manufacturer", "ModelNumber", "WarrantyGuarantorParts", "WarrantyDurationParts", "ReplacementCost", "ExpectedLife", "Material", "Finish"],
        "Component": ["Name", "TypeName", "Space", "SerialNumber", "InstallationDate", "WarrantyStartDate", "TagNumber", "AssetIdentifier"],
        "System": ["Name", "Category", "ComponentNames", "Description"],
        "Job": ["Name", "TypeName", "Frequency", "FrequencyUnit", "Duration", "DurationUnit", "Description"],
        "Resource": ["Name", "Category", "JobNames", "Description"],
        "Spare": ["Name", "TypeName", "Category", "Suppliers", "PartNumber", "Description"],
        "Contact": ["Name", "Email", "Company", "Phone", "Department", "Country"],
        "Document": ["Name", "Stage", "SheetName", "RowName", "Directory", "File", "Description"],
        "Coordinate": ["Name", "SheetName", "RowName", "CoordinateXAxis", "CoordinateYAxis", "CoordinateZAxis"],
        "Attribute": ["Name", "SheetName", "RowName", "Value", "Unit", "Description"],
        "Connection": ["Name", "SheetName", "RowName1", "RowName2", "ConnectionType", "Description"],
        "Issue": ["Name", "SheetName", "RowName", "Risk", "Chance", "Impact", "Owner", "Mitigation"],
    }

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, cols in sheets.items():
            pd.DataFrame(columns=cols).to_excel(writer, sheet_name=name, index=False)

    buffer.seek(0)
    return buffer.read()
