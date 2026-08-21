from os.path import join, getsize
from openpyxl import load_workbook
from openpyxl.styles import Font
import shutil
import tempfile
from .helpers import upload_metadata_file, get_template_path


from .constants import METADATA_UPLOAD_PS_PATH, TEMPLATE_PATH
from .excel_utils import rename_headers, excel_columns
from ...utils import validate_schema


### Create submission file
def create_excel(soda, upload_boolean, local_destination):
    """
    Create an Excel file for submission metadata.

    Args:
        soda (dict): The soda object containing dataset metadata.
        upload_boolean (bool): Whether to upload the file to Pennsieve.
        destination_path (str): The path to save the Excel file.

    Returns:
        dict: A dictionary containing the size of the metadata file.
    """

    validate_schema(soda["dataset_metadata"]["submission"], "submission_schema.json")

    font_submission = Font(name="Calibri", size=14, bold=False)

    source = get_template_path("submission.xlsx")

    destination = join(METADATA_UPLOAD_PS_PATH, "submission.xlsx") if upload_boolean else local_destination

    try:
        shutil.copyfile(source, destination)
    except FileNotFoundError as e:
        raise e
    
    #TODO: Do not use an array for the non-array values; zipping for the sake of the ascii value is not necessary until milestone_achieved
    submission_metadata_list = [
        soda["dataset_metadata"]["submission"]
    ]

    # write to excel file
    wb = load_workbook(destination)
    ws1 = wb["Sheet1"]
    start_index = 2
    for column, submission_data in zip(excel_columns(start_index), submission_metadata_list):
        ws1[column + "2"] = submission_data["consortium_data_standard"]
        ws1[column + "3"] = submission_data["funding_consortium"]
        ws1[column + "4"] = submission_data["award_number"]


        milestone_achieved = submission_data["milestone_achieved"]
        if isinstance(milestone_achieved, str):
            milestone_achieved = [milestone_achieved]
        for col, milestone in zip(excel_columns(start_index), milestone_achieved):
            ws1[col + str(5)] = milestone
            
        completion_dates = submission_data["milestone_completion_date"]
        if isinstance(completion_dates, str):
            completion_dates = [completion_dates]
        for col, milestone_date in zip(excel_columns(start_index), completion_dates):
            ws1[col + str(6)] = milestone_date
        ws1[column + "2"].font = font_submission
        ws1[column + "3"].font = font_submission
        ws1[column + "4"].font = font_submission
        ws1[column + "5"].font = font_submission
        ws1[column + "6"].font = font_submission

    # TODO: should milestone completion date also be an array?
    range_extend_len = max(len(submission_metadata_list[0]["milestone_achieved"]), len(submission_metadata_list[0]["milestone_completion_date"]))
    rename_headers(ws1, range_extend_len, 2)

    wb.save(destination)

    print("Excel file created successfully at:", destination)

    wb.close()

    # calculate the size of the metadata file
    size = getsize(destination)

    

    ## if generating directly on Pennsieve, then call upload function and then delete the destination path
    if upload_boolean:
        print("Uploading Excel file to Pennsieve...")
        upload_metadata_file("submission.xlsx", soda, destination, True)
        print("Excel file uploaded successfully to Pennsieve.")
    return {"size": size}


soda = {
    "guided-options": {},
    "curation-mode": "guided",
    "ps-account-selected": {},
    "dataset-structure": {
        "folders": {
            "data": {
                "folders": {
                    "bucketing-data": {
                        "type": "local",
                        "files": {},
                        "folders": {
                            "code-folder": {
                                "type": "local",
                                "files": {},
                                "folders": {},
                                "action": [
                                    "new"
                                ],
                                "relativePath": "data/bucketing-data/code-folder/"
                            },
                            "docs-folder": {
                                "type": "local",
                                "files": {},
                                "folders": {},
                                "action": [
                                    "new"
                                ],
                                "relativePath": "data/bucketing-data/docs-folder/"
                            },
                            "fromsam1-data": {
                                "type": "local",
                                "files": {},
                                "folders": {},
                                "action": [
                                    "new"
                                ],
                                "relativePath": "data/bucketing-data/fromsam1-data/"
                            },
                            "fromsub1-data": {
                                "type": "local",
                                "files": {},
                                "folders": {},
                                "action": [
                                    "new"
                                ],
                                "relativePath": "data/bucketing-data/fromsub1-data/"
                            },
                            "protocol-folder": {
                                "type": "local",
                                "files": {},
                                "folders": {},
                                "action": [
                                    "new"
                                ],
                                "relativePath": "data/bucketing-data/protocol-folder/"
                            },
                            "sam-1-data": {
                                "type": "local",
                                "files": {},
                                "folders": {},
                                "action": [
                                    "new"
                                ],
                                "relativePath": "data/bucketing-data/sam-1-data/"
                            },
                            "sam-derived-data": {
                                "type": "local",
                                "files": {},
                                "folders": {},
                                "action": [
                                    "new"
                                ],
                                "relativePath": "data/bucketing-data/sam-derived-data/"
                            },
                            "sub-1-data": {
                                "type": "local",
                                "files": {},
                                "folders": {},
                                "action": [
                                    "new"
                                ],
                                "relativePath": "data/bucketing-data/sub-1-data/"
                            }
                        },
                        "action": [
                            "new"
                        ],
                        "relativePath": "data/bucketing-data/"
                    }
                },
                "files": {},
                "type": "virtual",
                "action": [
                    "new"
                ],
                "location": "local",
                "relativePath": "data/"
            },
            "primary": {
                "folders": {
                    "bucketing-data": {
                        "folders": {
                            "code-folder": {
                                "folders": {},
                                "files": {
                                    "inside-code.txt": {
                                        "path": "/Users/aaronm/Downloads/bucketing-data/code-folder/inside-code.txt",
                                        "location": "local",
                                        "description": "",
                                        "additional-metadata": "",
                                        "action": [
                                            "new"
                                        ],
                                        "extension": ".txt",
                                        "relativePath": "data/bucketing-data/code-folder/inside-code.txt"
                                    }
                                },
                                "type": "virtual",
                                "action": [
                                    "new"
                                ],
                                "location": "local"
                            },
                            "docs-folder": {
                                "folders": {},
                                "files": {
                                    "inside-docs.txt": {
                                        "path": "/Users/aaronm/Downloads/bucketing-data/docs-folder/inside-docs.txt",
                                        "location": "local",
                                        "description": "",
                                        "additional-metadata": "",
                                        "action": [
                                            "new"
                                        ],
                                        "extension": ".txt",
                                        "relativePath": "data/bucketing-data/docs-folder/inside-docs.txt"
                                    }
                                },
                                "type": "virtual",
                                "action": [
                                    "new"
                                ],
                                "location": "local"
                            },
                            "fromsam1-data": {
                                "folders": {},
                                "files": {
                                    "fromsam1-data-file.txt": {
                                        "path": "/Users/aaronm/Downloads/bucketing-data/fromsam1-data/fromsam1-data-file.txt",
                                        "location": "local",
                                        "description": "",
                                        "additional-metadata": "",
                                        "action": [
                                            "new"
                                        ],
                                        "extension": ".txt",
                                        "relativePath": "data/bucketing-data/fromsam1-data/fromsam1-data-file.txt"
                                    }
                                },
                                "type": "virtual",
                                "action": [
                                    "new"
                                ],
                                "location": "local"
                            },
                            "fromsub1-data": {
                                "folders": {},
                                "files": {
                                    "fromsub1-data-file.txt": {
                                        "path": "/Users/aaronm/Downloads/bucketing-data/fromsub1-data/fromsub1-data-file.txt",
                                        "location": "local",
                                        "description": "",
                                        "additional-metadata": "",
                                        "action": [
                                            "new"
                                        ],
                                        "extension": ".txt",
                                        "relativePath": "data/bucketing-data/fromsub1-data/fromsub1-data-file.txt"
                                    }
                                },
                                "type": "virtual",
                                "action": [
                                    "new"
                                ],
                                "location": "local"
                            },
                            "protocol-folder": {
                                "folders": {},
                                "files": {
                                    "inside-protocol.txt": {
                                        "path": "/Users/aaronm/Downloads/bucketing-data/protocol-folder/inside-protocol.txt",
                                        "location": "local",
                                        "description": "",
                                        "additional-metadata": "",
                                        "action": [
                                            "new"
                                        ],
                                        "extension": ".txt",
                                        "relativePath": "data/bucketing-data/protocol-folder/inside-protocol.txt"
                                    }
                                },
                                "type": "virtual",
                                "action": [
                                    "new"
                                ],
                                "location": "local"
                            },
                            "sam-1-data": {
                                "folders": {},
                                "files": {
                                    "sam-1-data-file.txt": {
                                        "path": "/Users/aaronm/Downloads/bucketing-data/sam-1-data/sam-1-data-file.txt",
                                        "location": "local",
                                        "description": "",
                                        "additional-metadata": "",
                                        "action": [
                                            "new"
                                        ],
                                        "extension": ".txt",
                                        "relativePath": "data/bucketing-data/sam-1-data/sam-1-data-file.txt"
                                    }
                                },
                                "type": "virtual",
                                "action": [
                                    "new"
                                ],
                                "location": "local"
                            },
                            "sam-derived-data": {
                                "folders": {},
                                "files": {
                                    "sam-derived-data-file.txt": {
                                        "path": "/Users/aaronm/Downloads/bucketing-data/sam-derived-data/sam-derived-data-file.txt",
                                        "location": "local",
                                        "description": "",
                                        "additional-metadata": "",
                                        "action": [
                                            "new"
                                        ],
                                        "extension": ".txt",
                                        "relativePath": "data/bucketing-data/sam-derived-data/sam-derived-data-file.txt"
                                    }
                                },
                                "type": "virtual",
                                "action": [
                                    "new"
                                ],
                                "location": "local"
                            },
                            "sub-1-data": {
                                "folders": {},
                                "files": {
                                    "sub-1-data-file.txt": {
                                        "path": "/Users/aaronm/Downloads/bucketing-data/sub-1-data/sub-1-data-file.txt",
                                        "location": "local",
                                        "description": "",
                                        "additional-metadata": "",
                                        "action": [
                                            "new"
                                        ],
                                        "extension": ".txt",
                                        "relativePath": "data/bucketing-data/sub-1-data/sub-1-data-file.txt"
                                    }
                                },
                                "type": "virtual",
                                "action": [
                                    "new"
                                ],
                                "location": "local"
                            }
                        },
                        "files": {
                            "code-file.txt": {
                                "path": "/Users/aaronm/Downloads/bucketing-data/code-file.txt",
                                "location": "local",
                                "description": "",
                                "additional-metadata": "",
                                "action": [
                                    "new"
                                ],
                                "extension": ".txt",
                                "relativePath": "data/bucketing-data/code-file.txt"
                            },
                            "docs-file.txt": {
                                "path": "/Users/aaronm/Downloads/bucketing-data/docs-file.txt",
                                "location": "local",
                                "description": "",
                                "additional-metadata": "",
                                "action": [
                                    "new"
                                ],
                                "extension": ".txt",
                                "relativePath": "data/bucketing-data/docs-file.txt"
                            },
                            "fromsam1-data.txt": {
                                "path": "/Users/aaronm/Downloads/bucketing-data/fromsam1-data.txt",
                                "location": "local",
                                "description": "",
                                "additional-metadata": "",
                                "action": [
                                    "new"
                                ],
                                "extension": ".txt",
                                "relativePath": "data/bucketing-data/fromsam1-data.txt"
                            },
                            "fromsub1-data.txt": {
                                "path": "/Users/aaronm/Downloads/bucketing-data/fromsub1-data.txt",
                                "location": "local",
                                "description": "",
                                "additional-metadata": "",
                                "action": [
                                    "new"
                                ],
                                "extension": ".txt",
                                "relativePath": "data/bucketing-data/fromsub1-data.txt"
                            },
                            "protocol-file.txt": {
                                "path": "/Users/aaronm/Downloads/bucketing-data/protocol-file.txt",
                                "location": "local",
                                "description": "",
                                "additional-metadata": "",
                                "action": [
                                    "new"
                                ],
                                "extension": ".txt",
                                "relativePath": "data/bucketing-data/protocol-file.txt"
                            },
                            "sam-1-data-file.txt": {
                                "path": "/Users/aaronm/Downloads/bucketing-data/sam-1-data-file.txt",
                                "location": "local",
                                "description": "",
                                "additional-metadata": "",
                                "action": [
                                    "new"
                                ],
                                "extension": ".txt",
                                "relativePath": "data/bucketing-data/sam-1-data-file.txt"
                            },
                            "sam-derived-data.txt": {
                                "path": "/Users/aaronm/Downloads/bucketing-data/sam-derived-data.txt",
                                "location": "local",
                                "description": "",
                                "additional-metadata": "",
                                "action": [
                                    "new"
                                ],
                                "extension": ".txt",
                                "relativePath": "data/bucketing-data/sam-derived-data.txt"
                            },
                            "soda-for-sparc-19.0.0-beta.dmg": {
                                "path": "/Users/aaronm/Downloads/bucketing-data/soda-for-sparc-19.0.0-beta.dmg",
                                "location": "local",
                                "description": "",
                                "additional-metadata": "",
                                "action": [
                                    "new"
                                ],
                                "extension": ".dmg",
                                "relativePath": "data/bucketing-data/soda-for-sparc-19.0.0-beta.dmg"
                            },
                            "sub-1-data-file.txt": {
                                "path": "/Users/aaronm/Downloads/bucketing-data/sub-1-data-file.txt",
                                "location": "local",
                                "description": "",
                                "additional-metadata": "",
                                "action": [
                                    "new"
                                ],
                                "extension": ".txt",
                                "relativePath": "data/bucketing-data/sub-1-data-file.txt"
                            }
                        },
                        "type": "virtual",
                        "action": [
                            "new"
                        ],
                        "location": "local"
                    }
                },
                "files": {},
                "type": "virtual",
                "action": [
                    "new"
                ],
                "location": "local"
            }
        },
        "files": {},
        "relativePath": "/"
    },
    "generate-dataset": {
        "destination": "ps"
    },
    "starting-point": {
        "origin": "new"
    },
    "dataset_metadata": {
        "manifest_file": [
            {
                "filename": "primary",
                "timestamp": "",
                "description": "",
                "file_type": "folder",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data",
                "timestamp": "",
                "description": "",
                "file_type": "folder",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/code-file.txt",
                "timestamp": "2026-07-22T18:20:30Z",
                "description": "",
                "file_type": ".txt",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/code-folder",
                "timestamp": "",
                "description": "",
                "file_type": "folder",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/code-folder/inside-code.txt",
                "timestamp": "2026-07-22T18:20:30Z",
                "description": "",
                "file_type": ".txt",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/docs-file.txt",
                "timestamp": "2026-07-22T18:20:30Z",
                "description": "",
                "file_type": ".txt",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/docs-folder",
                "timestamp": "",
                "description": "",
                "file_type": "folder",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/docs-folder/inside-docs.txt",
                "timestamp": "2026-07-22T18:20:30Z",
                "description": "",
                "file_type": ".txt",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/fromsam1-data",
                "timestamp": "",
                "description": "",
                "file_type": "folder",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/fromsam1-data.txt",
                "timestamp": "2026-07-22T18:20:30Z",
                "description": "",
                "file_type": ".txt",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/fromsam1-data/fromsam1-data-file.txt",
                "timestamp": "2026-07-22T18:20:30Z",
                "description": "",
                "file_type": ".txt",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/fromsub1-data",
                "timestamp": "",
                "description": "",
                "file_type": "folder",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/fromsub1-data.txt",
                "timestamp": "2026-07-22T18:20:30Z",
                "description": "",
                "file_type": ".txt",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/fromsub1-data/fromsub1-data-file.txt",
                "timestamp": "2026-07-22T18:20:30Z",
                "description": "",
                "file_type": ".txt",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/protocol-file.txt",
                "timestamp": "2026-07-22T18:20:30Z",
                "description": "",
                "file_type": ".txt",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/protocol-folder",
                "timestamp": "",
                "description": "",
                "file_type": "folder",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/protocol-folder/inside-protocol.txt",
                "timestamp": "2026-07-22T18:20:30Z",
                "description": "",
                "file_type": ".txt",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/sam-1-data",
                "timestamp": "",
                "description": "",
                "file_type": "folder",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/sam-1-data-file.txt",
                "timestamp": "2026-07-22T18:20:30Z",
                "description": "",
                "file_type": ".txt",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/sam-1-data/sam-1-data-file.txt",
                "timestamp": "2026-07-22T18:20:30Z",
                "description": "",
                "file_type": ".txt",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/sam-derived-data",
                "timestamp": "",
                "description": "",
                "file_type": "folder",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/sam-derived-data.txt",
                "timestamp": "2026-07-22T18:20:30Z",
                "description": "",
                "file_type": ".txt",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/sam-derived-data/sam-derived-data-file.txt",
                "timestamp": "2026-07-22T18:20:30Z",
                "description": "",
                "file_type": ".txt",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/soda-for-sparc-19.0.0-beta.dmg",
                "timestamp": "2026-08-18T16:13:29.631610Z",
                "description": "",
                "file_type": ".dmg",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/sub-1-data",
                "timestamp": "",
                "description": "",
                "file_type": "folder",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/sub-1-data-file.txt",
                "timestamp": "2026-07-22T18:20:30Z",
                "description": "",
                "file_type": ".txt",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            },
            {
                "filename": "primary/bucketing-data/sub-1-data/sub-1-data-file.txt",
                "timestamp": "2026-07-22T18:20:30Z",
                "description": "",
                "file_type": ".txt",
                "entity": "",
                "data_modality": "",
                "also_in_dataset": "",
                "data_dictionary_path": "",
                "entity_is_transitive": "",
                "additional_metadata": ""
            }
        ],
        "submission": {
            "consortium_data_standard": "SPARC",
            "funding_consortium": "SPARC",
            "award_number": "",
            "milestone_achieved": ["one", "two"],
            "milestone_completion_date": [
                "2026-08-11T07:00:00.000Z",
                "2026-08-21T07:00:00.000Z",
                "2026-08-23T07:00:00.000Z",
            ]
        },
        "dataset_description": {
            "metadata_version": "3.0.2",
            "dataset_type": "computational",
            "standards_information": [
                {
                    "data_standard": "SPARC",
                    "data_standard_version": "2025.05.01"
                },
                {
                    "data_standard": "SODA",
                    "data_standard_version": "18.2.1-beta"
                }
            ],
            "basic_information": {
                "title": "ffd",
                "subtitle": "fgdfd",
                "description": "",
                "keywords": [],
                "funding": [],
                "acknowledgments": "",
                "license": "CC-BY-4.0"
            },
            "funding_information": {
                "funding_consortium": "SPARC",
                "funding_agency": "NIH",
                "award_number": ""
            },
            "study_information": {
                "study_purpose": "",
                "study_data_collection": "",
                "study_primary_conclusion": "",
                "study_organ_system": [],
                "study_approach": [],
                "study_technique": [],
                "study_collection_title": ""
            },
            "contributor_information": [
                {
                    "contributor_name": "marr, chris",
                    "contributor_orcid_id": "https://orcid.org/0000-0000-0000-0001",
                    "contributor_affiliation": "https://ror.org/04ttjf776",
                    "contributor_roles": [
                        "PrincipalInvestigator"
                    ]
                }
            ],
            "related_resource_information": [],
            "participant_information": {
                "number_of_subjects": 0,
                "number_of_samples": 0,
                "number_of_sites": 0,
                "number_of_performances": 0
            }
        },
        "README.md": "ddadssad"
    },
    "dataset-type": "computational",
    "dataset_performances": [],
    "funding_agency": "NIH",
    "save-file-path": "/Users/aaronm/SODA/Guided-Progress/ffd-3bbe95d7.json",
    "path-to-local-dataset-copy": "/Users/aaronm/Downloads/ffd"
}



try:
    create_excel(soda, False, "submission.xlsx")
except Exception as e:
    print(e)

