import os
import requests
import pandas as pd
import itertools
from ...constants import PENNSIEVE_URL
from ...utils import create_request_headers, get_access_token, get_dataset_id, PennsieveActionNoPermission
from ..permissions import pennsieve_get_current_user_permissions
from ..metadata import load_metadata_to_dataframe

from .. import logger


METADATA_FILES_SPARC = [
        "submission.xlsx",
        "submission.csv",
        "submission.json",
        "dataset_description.xlsx",
        "dataset_description.csv",
        "dataset_description.json",
        "subjects.xlsx",
        "subjects.csv",
        "subjects.json",
        "samples.xlsx",
        "samples.csv",
        "samples.json",
        "README.txt",
        "CHANGES.txt",
        "code_description.xlsx",
        "inputs_metadata.xlsx",
        "outputs_metadata.xlsx",
        "manifest.xlsx",
        "manifest.csv"
    ]


def import_pennsieve_dataset(soda_json_structure, requested_sparc_only=True):
    global logger
    high_level_sparc_folders = [
        "code",
        "derivative",
        "docs",
        "primary",
        "protocol",
        "source",
    ]
    manifest_sparc = ["manifest.xlsx", "manifest.csv"]
    high_level_metadata_sparc = [
        "submission.xlsx",
        "submission.csv",
        "submission.json",
        "dataset_description.xlsx",
        "dataset_description.csv",
        "dataset_description.json",
        "subjects.xlsx",
        "subjects.csv",
        "subjects.json",
        "samples.xlsx",
        "samples.csv",
        "samples.json",
        "README.txt",
        "CHANGES.txt",
        "code_description.xlsx",
        "inputs_metadata.xlsx",
        "outputs_metadata.xlsx",
    ]
    double_extensions = [
        ".tar.gz",
        ".tar.bz2",
        ".tar.xz",
        ".tar.zst",
        ".tar.lz",
        ".tar.lz4",
        ".tar.lzma",
        ".tgz",
        ".tbz",
        ".tbz2",
        ".txz",
        ".csv.gz",
        ".tsv.gz",
        ".txt.gz",
        ".json.gz",
        ".jsonl.gz",
        ".ndjson.gz",
        ".xml.gz",
        ".yaml.gz",
        ".yml.gz",
        ".log.gz",
        ".dat.gz",
        ".data.gz",
        ".bin.gz",
        ".sql.gz",
        ".nii.gz",
        ".nii.z",
        ".mgh.gz",
        ".gii.gz",
        ".annot.gz",
        ".label.gz",
        ".trk.gz",
        ".tck.gz",
        ".mif.gz",
        ".fif.gz",
        ".edf.gz",
        ".bdf.gz",
        ".set.gz",
        ".fdt.gz",
        ".vhdr.gz",
        ".vmrk.gz",
        ".eeg.gz",
        ".dcm.gz",
        ".dicom.gz",
        ".ima.gz",
        ".nrrd.gz",
        ".nhdr.gz",
        ".mha.gz",
        ".mhd.gz",
        ".hdr.gz",
        ".img.gz",
        ".ome.tif",
        ".ome.tiff",
        ".ome.btf",
        ".ome.tif2",
        ".ome.tif8",
        ".ome.xml",
        ".ome.zarr",
        ".ome.zarr.zip",
        ".tif.gz",
        ".tiff.gz",
        ".btf.gz",
        ".brukertiff.gz",
        ".mefd.gz",
        ".moberg.gz",
        ".fastq.gz",
        ".fq.gz",
        ".fasta.gz",
        ".fa.gz",
        ".fna.gz",
        ".faa.gz",
        ".sam.gz",
        ".bam.gz",
        ".cram.gz",
        ".vcf.gz",
        ".bcf.gz",
        ".bed.gz",
        ".gff.gz",
        ".gff3.gz",
        ".gtf.gz",
        ".wig.gz",
        ".bedgraph.gz",
        ".bg.gz",
        ".bcl.gz",
        ".scf.gz",
        ".vcf.bgz",
        ".bed.bgz",
        ".bedgraph.bgz",
        ".gff.bgz",
        ".gff3.bgz",
        ".gtf.bgz",
        ".h5.gz",
        ".hdf.gz",
        ".hdf5.gz",
        ".nc.gz",
        ".netcdf.gz",
        ".mat.gz",
        ".rds.gz",
        ".rda.gz",
        ".pickle.gz",
        ".pkl.gz",
        ".shp.zip",
        ".geojson.gz",
        ".gpkg.gz",
        ".kml.gz",
        ".gpx.gz",
        ".las.gz",
        ".sqlite.gz",
        ".db.gz",
        ".mdb.gz",
        ".parquet.gz",
        ".feather.gz",
        ".arrow.gz",
        ".cdf.gz",
        ".fits.gz",
        ".fit.gz",
        ".raw.gz",
        ".pdf.gz",
        ".html.gz",
        ".htm.gz",
        ".md.gz",
        ".rst.gz",
        ".zip.gz",
        ".7z.gz",
        ".rar.gz",
    ]

    global create_soda_json_completed
    global create_soda_json_total_items
    global create_soda_json_progress
    create_soda_json_progress = 0
    create_soda_json_total_items = 0
    create_soda_json_completed = 0

    # ["extensions"] doesn't seem to be returned by the Pennsieve API anymore
    def verify_file_name(file_name, extension):
        global logger
        if extension == "":
            return file_name

        double_ext = False
        for ext in double_extensions:
            if file_name.find(ext) != -1:
                double_ext = True
                break
                
        extension_from_name = ""

        if double_ext == False:
            extension_from_name = os.path.splitext(file_name)[1]
        else:
            extension_from_name = (
                os.path.splitext(os.path.splitext(file_name)[0])[1]
                + os.path.splitext(file_name)[1]
            )

        if extension_from_name == ("." + extension):
            return file_name
        else:
            return file_name + ("." + extension)


    def createFolderStructure(subfolder_json, manifest):
        """
            Function for creating the Pennsieve folder structure for a given dataset as an object stored locally.
            Arguments:
                subfolder_json: The json object containing the folder structure of the dataset
                pennsieve_client: The Pennsieve client object
                manifest: The manifest object for the dataset
        """
        # root level folder will pass subfolders into this function and will recursively check if there are subfolders while creating the json structure
        global logger
        global create_soda_json_progress
        
        collection_id = subfolder_json["path"]

        limit = 100
        offset = 0
        subfolder = []
        while True: 
            r = requests.get(f"{PENNSIEVE_URL}/packages/{collection_id}?limit={limit}&offset={offset}", headers=create_request_headers(get_access_token()))
            r.raise_for_status()
            page = r.json()["children"]
            subfolder.extend(page)
            if len(page) < limit:
                break
            offset += limit

        for items in subfolder:
            folder_item_name = items["content"]["name"]
            create_soda_json_progress += 1
            item_id = items["content"]["id"]
            # is a file name check if there are additional manifest information to attach to files
            if item_id[2:9] == "package":
                if (
                    folder_item_name[0:8] != "manifest"
                ):  # manifest files are not being included in json structure

                    # verify file name first (used for legacy Pennsieve datasets)
                    if("extension" not in subfolder):
                        folder_item_name = verify_file_name(folder_item_name, "")
                    else:
                        folder_item_name = verify_file_name(folder_item_name, subfolder["extension"])
                        
                    # verify timestamps
                    timestamp = items["content"]["createdAt"].replace('.', ',')

                    paths_list = [*subfolder_json["pspath"]]
                    subfolder_json["files"][folder_item_name] = {
                        "action": ["existing"],
                        "path": item_id,
                        "pspath": paths_list,
                        "timestamp": timestamp,
                        "location": "ps",
                        "additional-metadata": "",
                        "description": "",
                    }
                    
                    # creates path for folder_item_name (stored in temp_name)
                    if len(subfolder_json["files"][folder_item_name]["pspath"]) > 1:
                        temp_name = '/'.join(subfolder_json["files"][folder_item_name]["pspath"][1:]) + "/" + folder_item_name
                    else:
                        temp_name = folder_item_name
                    
                    if len(manifest.keys()) > 0:
                        # Dictionary that has the required manifest headers in lowercase and without spaces as keys
                        # and the correct manifest headers as values
                        defaultManifestHeadersNameMapped = {
                            "filename": "filename",
                            "timestamp": "timestamp",
                            "description": "description",
                            "filetype": "file type",
                            "entity": "entity",
                            "datamodality": "data modality",
                            "alsoindataset": "also in dataset",
                            "alsoindatasetpath": "also in dataset path",
                            "datadictionarypath": "data dictionary path",
                            "entityistransitive": "entity is transitive",
                            "additionalmetadata": "additional-metadata",
                        }

                        # Dictionary that will be used to store the correct manifest headers as keys
                        # and the values from the manifest as values
                        updated_manifest = {}

                        # Go through the imported manifest keys and change the keys to the correct name
                        # For example if the key is "File Name" change it to "filename"
                        for manifestKey in manifest.keys():
                            # Make the key lowercase
                            sterilizedKeyName = manifestKey.lower().replace(" ", "")
                            if sterilizedKeyName in defaultManifestHeadersNameMapped.keys():
                                # change the key to the correct name
                                # For example if the key name is "filetype" change it to "file type"
                                newManifestKeyName = defaultManifestHeadersNameMapped[sterilizedKeyName]
                                # Add the new key/value to the updated manifest
                                updated_manifest[newManifestKeyName] = manifest[manifestKey]
                            else:
                                # Keep the key/value the same and add it to the updated manifest
                                updated_manifest[manifestKey] = manifest[manifestKey]

                        if "filename" in updated_manifest.keys():
                            for manifestKey in updated_manifest.keys():
                                location_index = ""
                                # get the index of the file name in the manifest
                                if (temp_name in updated_manifest["filename"].values()):
                                    location_index = list(updated_manifest["filename"].values()).index(
                                        temp_name
                                    )
                                # This is for the case where the file name in the manifest has a slash at the beginning
                                # which is the case for files in the root folder
                                elif ("/" + temp_name in updated_manifest["filename"].values()):
                                    location_index = list(updated_manifest["filename"].values()).index(
                                        "/" + temp_name
                                    )
                                else:
                                    # break out of the for loop if the file name is not in the manifest
                                    break

                                # check if the key is in the required manifest headers, if it is, update the folder_item_name value
                                # corresponding to the key
                                if manifestKey in defaultManifestHeadersNameMapped.values():
                                    if updated_manifest[manifestKey][location_index] != "":
                                        if folder_item_name[0:1] == "/":
                                            subfolder_json["files"][folder_item_name[:1]][manifestKey] = updated_manifest[manifestKey][location_index]
                                        else:
                                            subfolder_json["files"][folder_item_name][manifestKey] = updated_manifest[manifestKey][location_index]
                                # if the key is not in the required manifest headers, add it to the extra columns folder_item_name value
                                else :
                                    # if the extra columns key does not exist, create it
                                    if "extra_columns" not in subfolder_json["files"][folder_item_name]:
                                        subfolder_json["files"][folder_item_name]["extra_columns"] = {}
                                    
                                    if updated_manifest[manifestKey][location_index] != "":
                                        subfolder_json["files"][folder_item_name]["extra_columns"][manifestKey] = updated_manifest[manifestKey][location_index]
                                    else:
                                        subfolder_json["files"][folder_item_name]["extra_columns"][manifestKey] = ""
                        else:
                            # filename not in updated manifest so recreate standard headers if they don't exist
                            # loop through the updated manifest keys and if header matches standard header add content else recreate
                            if len(updated_manifest.keys()) > 0:
                                location_index = ""
                                for manifestKey in updated_manifest.keys():
                                    if temp_name in updated_manifest[manifestKey].values():
                                        # file_names found
                                        location_index = list(updated_manifest[manifestKey].values()).index(
                                        temp_name
                                        )
                                    if ("/" + temp_name in updated_manifest[manifestKey].values()):
                                        location_index = list(updated_manifest[manifestKey].values()).index(
                                        "/" + temp_name
                                        )
                                    if location_index != "":
                                        if manifestKey in defaultManifestHeadersNameMapped.values():
                                            if folder_item_name[0:1] == "/":
                                                subfolder_json["files"][folder_item_name[1:]][manifestKey] = updated_manifest[manifestKey][location_index]
                                            else:
                                                subfolder_json["files"][folder_item_name][manifestKey] = updated_manifest[manifestKey][location_index]
                                        else:
                                            if "extra_columns" not in subfolder_json["files"][folder_item_name]:
                                                subfolder_json["files"][folder_item_name]["extra_columns"] = {}
                                            subfolder_json["files"][folder_item_name]["extra_columns"][manifestKey] = updated_manifest[manifestKey][location_index]

            else:  # another subfolder found
                paths_list = [*subfolder_json["pspath"], folder_item_name]
                subfolder_json["folders"][folder_item_name] = {
                    "action": ["existing"],
                    "path": item_id,
                    "pspath": paths_list,
                    "files": {},
                    "folders": {},
                    "location": "ps",
                }

        if len(subfolder_json["folders"].keys()) != 0:  # there are subfolders
            for folder in subfolder_json["folders"].keys():
                subfolder = subfolder_json["folders"][folder]
                createFolderStructure(subfolder, manifest)


    # check that the Pennsieve dataset is valid
    try:
        bf_dataset_name = soda_json_structure["ps-dataset-selected"]["dataset-name"]
    except Exception as e:
        raise e

    selected_dataset_id = get_dataset_id(bf_dataset_name)

    # check that the user has permission to edit this dataset
    role = pennsieve_get_current_user_permissions(selected_dataset_id, get_access_token())["role"]
    if role not in ["owner", "manager", "editor"]:
        curatestatus = "Done"
        raise PennsieveActionNoPermission("You do not have permissions to edit upload this Pennsieve dataset.")


    # surface layer of dataset is pulled. then go through through the children to get information on subfolders
    manifest_dict = {}
    manifest_error_message = []
    soda_json_structure["dataset-structure"] = {
        "files": {},
        "folders": {},
    }

    # root of dataset is pulled here (high level folders/files are gathered here)
    # root_folder is the files and folders within root
    r = requests.get(f"{PENNSIEVE_URL}/datasets/{selected_dataset_id}", headers=create_request_headers(get_access_token()))
    r.raise_for_status()
    root_folder = r.json()["children"]

    # Get the amount of files/folders in the dataset
    r = requests.get(f"{PENNSIEVE_URL}/datasets/{selected_dataset_id}/packageTypeCounts", headers=create_request_headers(get_access_token()))
    r.raise_for_status()
    packages_list = r.json()


    # root's children files
    for count in packages_list.values():
        create_soda_json_total_items += int(count)

    # set manifest dictionry to empty dictionary; used to store the manifest information while we import dataset
    manifest_dict = {}
    

    # Gather metadata files first
    for items in root_folder:
        item_id = items["content"]["id"]
        item_name = items["content"]["name"]

        # Import manifest at the root of the dataset
        if item_name in manifest_sparc:
            # Item is a manifest file
            df = ""
            try:
                if item_name.lower() == "manifest.xlsx":
                    df = load_metadata_to_dataframe(item_id, "excel", get_access_token())
                    df = df.fillna("")
                else:
                    df = load_metadata_to_dataframe(item_id, "csv", get_access_token())
                    df = df.fillna("")
                manifest_dict = df.to_dict()
            except Exception as e:
                manifest_error_message.append(item_name)

        # Item is a metadata file
        if item_name in high_level_metadata_sparc:
            create_soda_json_progress += 1
            if "dataset_metadata" not in soda_json_structure.keys():
                soda_json_structure["dataset_metadata"] = {}
            soda_json_structure["dataset_metadata"][item_name] = {
                "location": "ps",
                "action": ["existing"],
                "path": item_id,
            }

    # Process the folder structure
    for items in root_folder:
        item_id = items["content"]["id"]
        item_name = items["content"]["name"]

        # If package type is Collection, then it is a folder
        if items["content"]["packageType"] == "Collection" and item_name in high_level_sparc_folders:
            create_soda_json_progress += 1
            soda_json_structure["dataset-structure"]["folders"][item_name] = {
                "location": "ps",
                "path": item_id,
                "action": ["existing"],
                "files": {},
                "folders": {},
                "pspath": [item_name],
            }

            # Check the content of the folder to see if a manifest file exists
            r = requests.get(f"{PENNSIEVE_URL}/packages/{item_id}", headers=create_request_headers(get_access_token()))
            r.raise_for_status()
            folder_content = r.json()["children"]

            if len(folder_content) > 0:
                high_lvl_folder_dict = soda_json_structure["dataset-structure"]["folders"][item_name]

                createFolderStructure(
                    high_lvl_folder_dict, manifest_dict
                )  # Passing item's JSON and the collection ID

    success_message = (
        "Data files under a valid high-level SPARC folders have been imported"
    )
    create_soda_json_completed = 1

    logger.info(f"Time to import {soda_json_structure['ps-dataset-selected']['dataset-name']} ")
    return {
        "soda_object": soda_json_structure,
        "success_message": success_message,
        "manifest_error_message": manifest_error_message,
        "import_progress": create_soda_json_progress,
        "import_total_items": create_soda_json_total_items,
    }


create_soda_json_progress = 0
create_soda_json_total_items = 0
create_soda_json_completed = 0


def create_soda_json_object_backend(
    soda_json_structure, root_folder_path, irregularFolders, replaced
):
    """
    This function is meant for importing local datasets into SODA.
    It creates a json object with the structure of the dataset.
    """
    global create_soda_json_progress  # amount of items counted during recursion
    global create_soda_json_total_items  # counts the total items in folder
    global create_soda_json_completed  # completed progress is either 0 or 1
    global METADATA_FILES_SPARC

    high_level_sparc_folders = [
        "code",
        "derivative",
        "docs",
        "primary",
        "protocol",
        "source",
    ]

    dataset_folder = soda_json_structure["dataset-structure"] = {"folders": {}}

    def recursive_structure_create(dataset_structure, folder_path, root_manifest):
        global create_soda_json_progress
        # going within high level folders
        # add manifest details if manifest exists
        manifest_object = {
            "filename": "",
            "timestamp": "",
            "description": "",
            "file type": "",
            "entity": "",
            "data modality": "",
            "also in dataset": "",
            "also in dataset path": "",
            "data dictionary path": "",
            "entity is transitive": "",
            "additional-metadata": "",
        }

        lastSlash = folder_path.rfind("/") + 1
        folder_name = folder_path[lastSlash:]

        if folder_name in replaced.keys():
            folder_name = replaced[folder_name]

        # Check if folder is in irregular folders
        if folder_path in irregularFolders:
            index_check = irregularFolders.index(folder_path)
            modified_name = replaced[os.path.basename(folder_path)]
            folder_path = irregularFolders[index_check]


        entries = os.listdir(folder_path)
        for entry in entries:
            item_path = os.path.normpath(os.path.join(folder_path, entry))
            if os.path.isfile(item_path):
                # Check manifest to add metadata
                if entry[0:1] != "." and entry[0:8] != "manifest":
                    create_soda_json_progress += 1
                    # Use the root manifest to find metadata for the file
                    for row in root_manifest:
                        extra_columns = False
                        if len(row) > 11:
                            extra_columns = True
                            extra_columns_dict = dict(itertools.islice(row.items(), 5, len(row)))

                        if row["filename"] == entry:
                            # Add description metadata
                            manifest_object["description"] = row.get("description", "")
                            # Add additional metadata
                            manifest_object["additional-metadata"] = row.get("Additional Metadata", "")
                            if extra_columns:
                                manifest_object["extra_columns"] = extra_columns_dict

                    # Create JSON structure for the file
                    if "extra_columns" in manifest_object:
                        dataset_structure["files"][entry] = {
                            "path": item_path,
                            "location": "local",
                            "action": ["existing"],
                            "description": manifest_object["description"],
                            "additional-metadata": manifest_object["additional-metadata"],
                            "extra_columns": manifest_object["extra_columns"],
                        }
                    else:
                        dataset_structure["files"][entry] = {
                            "path": item_path,
                            "location": "local",
                            "action": ["existing"],
                            "description": manifest_object["description"],
                            "additional-metadata": manifest_object["additional-metadata"],
                        }
            elif os.path.isdir(item_path) is True:
                create_soda_json_progress += 1
                if item_path in irregularFolders:
                    index_check = irregularFolders.index(item_path)
                    modified_name = replaced[os.path.basename(item_path)]

                    dataset_structure["folders"][modified_name] = {
                        "folders": {},
                        "files": {},
                        "path": item_path,
                        "location": "local",
                        "action": ["existing"],
                        "original-name": entry,
                    }
                    for folder in dataset_structure["folders"][modified_name][
                        "folders"
                    ]:
                        updated_path = dataset_structure["folders"][modified_name][
                            folder
                        ]["path"]
                        recursive_structure_create(
                            dataset_structure["folders"][modified_name][folder],
                            updated_path,
                            root_manifest
                        )
                else:
                    dataset_structure["folders"][entry] = {
                        "folders": {},
                        "files": {},
                        "path": item_path,
                        "location": "local",
                        "action": ["existing"],
                    }

        for folder in dataset_structure["folders"]:
            updated_path = dataset_structure["folders"][folder]["path"]
            recursive_structure_create(
                dataset_structure["folders"][folder], updated_path, root_manifest
            )

    # BEGIN

    # Check for a single manifest file at the root of the dataset
    root_manifest_csv = os.path.join(root_folder_path, "manifest.csv")
    root_manifest_xlsx = os.path.join(root_folder_path, "manifest.xlsx")

    soda_json_structure["starting-point"]["manifest"] = {}

    if os.path.exists(root_manifest_csv):
        csv_data = pd.read_csv(root_manifest_csv)
        csv_data.fillna("", inplace=True)
        json_format = csv_data.to_dict(orient="records")
        soda_json_structure["starting-point"]["manifest"] = json_format
        soda_json_structure["starting-point"]["path"] = root_manifest_csv
    elif os.path.exists(root_manifest_xlsx):
        excel_data = pd.read_excel(root_manifest_xlsx, sheet_name="Sheet1")
        excel_data.fillna("", inplace=True)
        json_format = excel_data.to_dict(orient="records")
        soda_json_structure["starting-point"]["manifest"] = json_format
        soda_json_structure["starting-point"]["path"] = root_manifest_xlsx



    # count the amount of items in folder
    create_soda_json_total_items = 0
    for root, dirs, filenames in os.walk(root_folder_path):
        # walk through all folders and it's subfolders
        for Dir in dirs:
            # does not take hidden folders or manifest folders
            if Dir[0:1] != "." and Dir[0:8] != "manifest":
                create_soda_json_total_items += 1
        for fileName in filenames:
            if root == root_folder_path and fileName in METADATA_FILES_SPARC:
                # goes through all files and does not count hidden files
                create_soda_json_total_items += 1
            else:
                if fileName[0:1] != ".":
                    create_soda_json_total_items += 1

    # reading high level folders
    create_soda_json_completed = 0
    create_soda_json_progress = 0
    entries = os.listdir(root_folder_path)


    for entry in entries:
        # begin going through high level folders
        item_path = os.path.normpath(os.path.join(root_folder_path, entry))
        # high level folder paths
        if os.path.isfile(item_path) is True:
            if entry[0:1] != "." and entry in METADATA_FILES_SPARC:
                # is not a hidden folder
                create_soda_json_progress += 1
                soda_json_structure["dataset_metadata"][entry] = {
                    "path": item_path,
                    "location": "local",
                    "action": ["existing"],
                }
            # do file work here
        elif os.path.isdir(item_path) is True:
            create_soda_json_progress += 1
            # add item to soda
            if item_path in irregularFolders:
                index_check = irregularFolders.index(item_path)
                modified_name = replaced[index_check]
                folder_name = modified_name
                dataset_folder["folders"][folder_name] = {
                    "folders": {},
                    "files": {},
                    "path": item_path,
                    "location": "local",
                    "action": ["existing"],
                    "original-basename": item_path[(item_path.rfind("/") + 1) :],
                }
            else:
                if entry in high_level_sparc_folders:
                    dataset_folder["folders"][entry] = {
                        "folders": {},
                        "files": {},
                        "path": item_path,
                        "location": "local",
                        "action": ["existing"],
                    }
            soda_json_structure["starting-point"][entry] = {"path": ""}

    for folder in dataset_folder["folders"]:
        # go through high level folders again
        high_lvl_path = root_folder_path + "/" + folder
        recursive_structure_create(dataset_folder["folders"][folder], high_lvl_path, soda_json_structure["starting-point"]["manifest"])

    create_soda_json_completed = 1
    return soda_json_structure


def monitor_local_json_progress():
    """
    Function for monitoring progress of json_object_creation
    Used for progress bar
    """
    global create_soda_json_completed
    global create_soda_json_total_items
    global create_soda_json_progress
    progress_percentage = (
        create_soda_json_progress / create_soda_json_total_items
    ) * 100

    return {
        "create_soda_json_progress": create_soda_json_progress,
        "create_soda_json_total_items": create_soda_json_total_items,
        "progress_percentage": progress_percentage,
        "create_soda_json_completed": create_soda_json_completed
    }
