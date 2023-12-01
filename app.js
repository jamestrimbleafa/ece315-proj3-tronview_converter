document.getElementById('convertButton').addEventListener('click', () => {
    console.log("Converting");
    const fileInput = document.getElementById('excelFile');
    const file = fileInput.files[0];
    if (!file) {
        alert('Please select a file!');
        return;
    }

    const mode = document.getElementById('mode').value;
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = e.target.result;
        const workbook = XLSX.read(data, {type: 'binary'});
        
        // Select the sheet named "Analysis"
        const analysisSheet = workbook.Sheets["Analysis"];
        if (!analysisSheet) {
            alert('Sheet "Analysis" not found in the Excel file!');
            return;
        }

        // // Read the value from cell B3
        // const cell = analysisSheet['B12'];
        // if (!cell) {
        //     alert('Cell B3 is empty or does not exist in the "Analysis" sheet!');
        //     return;
        // }
        // const cellValue = cell.v; // Get the value from the cell
        // alert(cell.v); // Show the value in an alert box
        // // Double the value and create text file
        // createTextFile(cellValue * 2);
        if (mode == "conventional") {
            mpq26 = Math.ceil(analysisSheet['G36'].v);
            mpq30 = Math.ceil(analysisSheet['G37'].v);
            mpq32 = Math.ceil(analysisSheet['G38'].v);
        } else if (mode == "mixed") {
            mpq26 = Math.ceil(analysisSheet['H36'].v);
            mpq30 = Math.ceil(analysisSheet['H37'].v);
            mpq32 = Math.ceil(analysisSheet['H38'].v);
        } else { // stealth
            mpq26 = Math.ceil(analysisSheet['I36'].v);
            mpq30 = Math.ceil(analysisSheet['I37'].v);
            mpq32 = Math.ceil(analysisSheet['I38'].v);
        }

        const text = "{ \
            \"version\": 2.1, \
            \"name\": \"ECE315_Project3Map_"+(analysisSheet['A3'].v ?? '')+"\", \
            \"date\": \"12/1/2023\", \
            \"spi\": { \
                \"elevationAsyncStatus\": { \
                    \"currentRequestStatus\": \"fulfilled\", \
                    \"currentRequestId\": \"dyWtljWPMIJGTzDeKr9lN\" \
                }, \
                \"declinationAsyncStatus\": { \
                    \"currentRequestStatus\": \"fulfilled\", \
                    \"currentRequestId\": \"hr-VNVvq-cRTbSVUUFLsL\" \
                }, \
                \"position\": { \
                    \"lat\": 44.935676709405755, \
                    \"lng\": 140.67437293415352, \
                    \"declination\": -10.56003, \
                    \"elevation\": 0 \
                } \
            }, \
            \"bullseye\": { \
                \"modeSelected\": false, \
                \"elevationAsyncStatus\": { \
                    \"currentRequestStatus\": \"fulfilled\", \
                    \"currentRequestId\": \"\" \
                }, \
                \"declinationAsyncStatus\": { \
                    \"currentRequestStatus\": \"fulfilled\", \
                    \"currentRequestId\": \"_D0zLc4k_IMQ6mGn_AOeS\" \
                }, \
                \"degreesBetweenRadials\": 45, \
                \"numRings\": 6, \
                \"position\": { \
                    \"lat\": 44.77793589631623, \
                    \"lng\": 139.52087402343753, \
                    \"declination\": -10.70723, \
                    \"elevation\": 0 \
                }, \
                \"ringSpacing\": 20, \
                \"spacingUnits\": \"nm\", \
                \"title\": \"Bullseye\", \
                \"color\": \"#00f\" \
            }, \
            \"entities\": { \
                \"linkedLists\": [ \
                    { \
                        \"id\": \"95511454-0c1a-4111-aae6-3d8784fb453d\", \
                        \"category\": \"ROUTE\", \
                        \"name\": \"Conventional\", \
                        \"color\": \"#7ed321\" \
                    }, \
                    { \
                        \"id\": \"f3f222cb-e8ec-4caa-87d6-d28488d844bb\", \
                        \"category\": \"ROUTE\", \
                        \"name\": \"Mixed\", \
                        \"color\": \"#f8e71c\" \
                    }, \
                    { \
                        \"id\": \"0d3a42bf-9162-4661-8d7c-99866c118627\", \
                        \"category\": \"ROUTE\", \
                        \"name\": \"Stealth\", \
                        \"color\": \"#d0021b\" \
                    } \
                ], \
                \"steerpoints\": [ \
                    { \
                        \"position\": { \
                            \"lat\": 43.5411514, \
                            \"lng\": 142.1490453, \
                            \"declination\": -10, \
                            \"elevation\": 73 \
                        }, \
                        \"id\": \"ee281039-d447-42f0-acd0-8048f4410d7c\", \
                        \"title\": \"SAM\", \
                        \"declinationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"elevationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"label\": { \
                            \"position\": { \
                                \"lat\": 43.5411514, \
                                \"lng\": 142.1490453 \
                            }, \
                            \"offset\": \"left\", \
                            \"enabled\": true \
                        }, \
                        \"ring\": { \
                            \"radius\": 10, \
                            \"dashed\": false \
                        }, \
                        \"square\": { \
                            \"width\": 0, \
                            \"orientation\": 360 \
                        } \
                    }, \
                    { \
                        \"position\": { \
                            \"lat\": 42.3197772, \
                            \"lng\": 140.9940825, \
                            \"declination\": -10, \
                            \"elevation\": 10 \
                        }, \
                        \"id\": \"ee281039-d447-42f0-acd0-8048f2430d7c\", \
                        \"title\": \"MPQ-26\", \
                        \"declinationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"elevationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"label\": { \
                            \"position\": { \
                                \"lat\": 42.3197772, \
                                \"lng\": 140.9940825 \
                            }, \
                            \"offset\": \"left\", \
                            \"enabled\": true \
                        }, \
                        \"category\": \"THREAT\", \
                        \"icon\": \"EW Radar\", \
                        \"ring\": { \
                            \"radius\": "+mpq26+", \
                            \"dashed\": false \
                        }, \
                        \"square\": { \
                            \"width\": 0, \
                            \"orientation\": 360 \
                        } \
                    }, \
                    { \
                        \"position\": { \
                            \"lat\": 42.4506164, \
                            \"lng\": 143.2311628, \
                            \"declination\": -10, \
                            \"elevation\": 114 \
                        }, \
                        \"id\": \"ee281039-d447-42f0-acd0-8048f4330d7c\", \
                        \"title\": \"MPQ-26\", \
                        \"declinationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"elevationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"label\": { \
                            \"position\": { \
                                \"lat\": 42.4506164, \
                                \"lng\": 143.2311628 \
                            }, \
                            \"offset\": \"left\", \
                            \"enabled\": true \
                        }, \
                        \"category\": \"THREAT\", \
                        \"icon\": \"EW Radar\", \
                        \"ring\": { \
                            \"radius\": "+mpq26+", \
                            \"dashed\": false \
                        }, \
                        \"square\": { \
                            \"width\": 0, \
                            \"orientation\": 360 \
                        } \
                    }, \
                    { \
                        \"position\": { \
                            \"lat\": 43.0408994, \
                            \"lng\": 144.1930083, \
                            \"declination\": -10, \
                            \"elevation\": 93 \
                        }, \
                        \"id\": \"ee281039-d447-42f0-acd0-804844430d7c\", \
                        \"title\": \"MPQ-26\", \
                        \"declinationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"elevationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"label\": { \
                            \"position\": { \
                                \"lat\": 43.0408994, \
                                \"lng\": 144.1930083 \
                            }, \
                            \"offset\": \"left\", \
                            \"enabled\": true \
                        }, \
                        \"category\": \"THREAT\", \
                        \"icon\": \"EW Radar\", \
                        \"ring\": { \
                            \"radius\": "+mpq26+", \
                            \"dashed\": false \
                        }, \
                        \"square\": { \
                            \"width\": 0, \
                            \"orientation\": 360 \
                        } \
                    }, \
                    { \
                        \"position\": { \
                            \"lat\": 42.65496, \
                            \"lng\": 142.1929081, \
                            \"declination\": -10, \
                            \"elevation\": 154 \
                        }, \
                        \"id\": \"ee281039-d447-42f0-acd0-8048f4435d7c\", \
                        \"title\": \"MPQ-30\", \
                        \"declinationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"elevationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"label\": { \
                            \"position\": { \
                                \"lat\": 42.65496, \
                                \"lng\": 142.1929081 \
                            }, \
                            \"offset\": \"left\", \
                            \"enabled\": true \
                        }, \
                        \"category\": \"THREAT\", \
                        \"icon\": \"EW Radar\", \
                        \"ring\": { \
                            \"radius\": "+mpq30+", \
                            \"dashed\": false \
                        }, \
                        \"square\": { \
                            \"width\": 0, \
                            \"orientation\": 360 \
                        } \
                    }, \
                    { \
                        \"position\": { \
                            \"lat\": 42.9218147, \
                            \"lng\": 143.2097319, \
                            \"declination\": -10, \
                            \"elevation\": 39 \
                        }, \
                        \"id\": \"ee281039-d447-42f0-acd0-8048f4630d7c\", \
                        \"title\": \"MPQ-30\", \
                        \"declinationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"elevationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"label\": { \
                            \"position\": { \
                                \"lat\": 42.9218147, \
                                \"lng\": 143.2097319 \
                            }, \
                            \"offset\": \"left\", \
                            \"enabled\": true \
                        }, \
                        \"category\": \"THREAT\", \
                        \"icon\": \"EW Radar\", \
                        \"ring\": { \
                            \"radius\": "+mpq30+", \
                            \"dashed\": false \
                        }, \
                        \"square\": { \
                            \"width\": 0, \
                            \"orientation\": 360 \
                        } \
                    }, \
                    { \
                        \"position\": { \
                            \"lat\": 43.7477542, \
                            \"lng\": 143.4617647, \
                            \"declination\": -10, \
                            \"elevation\": 310 \
                        }, \
                        \"id\": \"ee281039-d447-42f0-acd0-8048f4730d7c\", \
                        \"title\": \"MPQ-30\", \
                        \"declinationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"elevationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"label\": { \
                            \"position\": { \
                                \"lat\": 43.7477542, \
                                \"lng\": 143.4617647 \
                            }, \
                            \"offset\": \"left\", \
                            \"enabled\": true \
                        }, \
                        \"category\": \"THREAT\", \
                        \"icon\": \"EW Radar\", \
                        \"ring\": { \
                            \"radius\": "+mpq30+", \
                            \"dashed\": false \
                        }, \
                        \"square\": { \
                            \"width\": 0, \
                            \"orientation\": 360 \
                        } \
                    }, \
                    { \
                        \"position\": { \
                            \"lat\": 43.3765122, \
                            \"lng\": 142.4110667, \
                            \"declination\": -10, \
                            \"elevation\": 171 \
                        }, \
                        \"id\": \"ee281039-d447-42f0-acd0-8048f4480d7c\", \
                        \"title\": \"MPQ-32\", \
                        \"declinationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"elevationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"label\": { \
                            \"position\": { \
                                \"lat\": 43.3765122, \
                                \"lng\": 142.4110667 \
                            }, \
                            \"offset\": \"left\", \
                            \"enabled\": true \
                        }, \
                        \"category\": \"THREAT\", \
                        \"icon\": \"EW Radar\", \
                        \"ring\": { \
                            \"radius\": "+mpq32+", \
                            \"dashed\": false \
                        }, \
                        \"square\": { \
                            \"width\": 0, \
                            \"orientation\": 360 \
                        } \
                    }, \
                    { \
                        \"position\": { \
                            \"lat\": 43.3072086, \
                            \"lng\": 141.7404672, \
                            \"declination\": -10, \
                            \"elevation\": 11 \
                        }, \
                        \"id\": \"ee281039-d447-42f0-acd0-8048f443097c\", \
                        \"title\": \"MPQ-32\", \
                        \"declinationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"elevationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"label\": { \
                            \"position\": { \
                                \"lat\": 43.3072086, \
                                \"lng\": 141.7404672 \
                            }, \
                            \"offset\": \"left\", \
                            \"enabled\": true \
                        }, \
                        \"category\": \"THREAT\", \
                        \"icon\": \"EW Radar\", \
                        \"ring\": { \
                            \"radius\": "+mpq32+", \
                            \"dashed\": false \
                        }, \
                        \"square\": { \
                            \"width\": 0, \
                            \"orientation\": 360 \
                        } \
                    }, \
                    { \
                        \"position\": { \
                            \"lat\": 44.1284403, \
                            \"lng\": 142.4992075, \
                            \"declination\": -10, \
                            \"elevation\": 174 \
                        }, \
                        \"id\": \"ee281039-d447-42f0-acd0-8048f1030d7c\", \
                        \"title\": \"MPQ-32\", \
                        \"declinationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"elevationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"label\": { \
                            \"position\": { \
                                \"lat\": 44.1284403, \
                                \"lng\": 142.4992075 \
                            }, \
                            \"offset\": \"left\", \
                            \"enabled\": true \
                        }, \
                        \"category\": \"THREAT\", \
                        \"icon\": \"EW Radar\", \
                        \"ring\": { \
                            \"radius\": "+mpq32+", \
                            \"dashed\": false \
                        }, \
                        \"square\": { \
                            \"width\": 0, \
                            \"orientation\": 360 \
                        } \
                    }, \
                    { \
                        \"position\": { \
                            \"lat\": 43.5252592, \
                            \"lng\": 143.1536442, \
                            \"declination\": -10, \
                            \"elevation\": 675 \
                        }, \
                        \"id\": \"ee281039-d447-42f0-acd0-8048f1130d7c\", \
                        \"title\": \"TGT-EAST:\", \
                        \"declinationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"elevationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"label\": { \
                            \"position\": { \
                                \"lat\": 43.5252592, \
                                \"lng\": 143.1536442 \
                            }, \
                            \"offset\": \"left\", \
                            \"enabled\": true \
                        }, \
                        \"category\": \"UNASSIGNED_ROUTE_POINT\", \
                        \"icon\": \"Target\", \
                        \"ring\": { \
                            \"radius\": 0, \
                            \"dashed\": false \
                        }, \
                        \"square\": { \
                            \"width\": 0, \
                            \"orientation\": 360 \
                        } \
                    }, \
                    { \
                        \"position\": { \
                            \"lat\": 43.1199461, \
                            \"lng\": 141.5813731, \
                            \"declination\": -10, \
                            \"elevation\": 0 \
                        }, \
                        \"id\": \"ee281039-d447-42f0-acd0-804812430d7c\", \
                        \"title\": \"TGT-WEST:\", \
                        \"declinationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"elevationAsyncStatus\": { \
                            \"currentRequestStatus\": \"fulfilled\", \
                            \"currentRequestId\": \"\" \
                        }, \
                        \"label\": { \
                            \"position\": { \
                                \"lat\": 43.1199461, \
                                \"lng\": 141.5813731 \
                            }, \
                            \"offset\": \"left\", \
                            \"enabled\": true \
                        }, \
                        \"category\": \"UNASSIGNED_ROUTE_POINT\", \
                        \"icon\": \"Target\", \
                        \"ring\": { \
                            \"radius\": 0, \
                            \"dashed\": false \
                        }, \
                        \"square\": { \
                            \"width\": 0, \
                            \"orientation\": 360 \
                        } \
                    } \
                ] \
            }, \
            \"ui\": { \
                \"commonUi\": { \
                    \"activeRoute\": \"noRoute\", \
                    \"activeLine\": \"noLine\" \
                }, \
                \"leafletUi\": { \
                    \"center\": { \
                        \"lat\": 42.928651253962215, \
                        \"lng\": 142.27651777064864 \
                    }, \
                    \"markerSize\": 3, \
                    \"zoom\": 7.5 \
                } \
            } \
        }";
        createTextFile(text);

    };
    
    reader.readAsBinaryString(file);
});

function createTextFile(value) {
    const textContent = String(value); // Convert the value to a string

    const modeDropdown = document.getElementById('mode');
    const selectedMode = modeDropdown.value; // Get the value of the mode dropdown input

    const blob = new Blob([textContent], {type: 'text/plain'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = selectedMode + '.txt'; // Use the value for the download filename
    a.click();
}
