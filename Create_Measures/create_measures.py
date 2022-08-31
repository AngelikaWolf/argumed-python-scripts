from asyncio.log import logger
import requests
import json
from requests.structures import CaseInsensitiveDict
import pandas as pd
from datetime import datetime
import locale

locale.setlocale(locale.LC_TIME, "de_DE")  # German

# How to use:
# The risk assessment file must be stored in the same location as this script.
# You also need a file for the Weissbier access token. The name must be "token.txt" and it must also be stored in the same location as this script.
# Then you can run the script.

DEV_URL = "https://weissbier.dev.dialoguecorp.com/api/v1/customers"
PROD_URL = "https://weissbier.dialoguecorp.eu/api/v1/customers"

# Headers for Request and Token
with open("token.txt", "r") as f:
    token = f.read()
    token = token.strip()

headers = CaseInsensitiveDict()
headers["Accept"] = "application/json"
headers["Authorization"] = f"Bearer {token}"
headers["Content-Type"] = "application/json"

while True:

    # Get the name of the file and check if it exists
    file_name = input("Please enter the name of the Excel file.\n")
    try:
        open(file_name, "r")
        break
    except FileNotFoundError as e:
        logger.error("The file does not exist. Please try again.")


while True:

    # Get the name of the sheet and check if it exists
    sheet_name = input(
        "Please enter the name of the sheet (tab) that contains the risk assessment.\n"
    )
    try:
        get_first_row = pd.read_excel(
            file_name, sheet_name=sheet_name, nrows=1, usecols="A"
        )
        break
    except ValueError as error:
        logger.error("The sheet does not exist. Please try again.")

# Open the file and then get the first row (this is needed for the source field)
get_first_row = pd.read_excel(file_name, sheet_name=sheet_name, nrows=1, usecols="A")
get_gbu_name = str(get_first_row)

# Get only the name of the risk assessment
gbu_name = get_gbu_name.partition("\n")[0]
gbu_name = " ".join(gbu_name.split())  # Remove extra Spaces

# Filter Rows of the Excel to extract the necessary information
Row_list = []
df = pd.read_excel(
    file_name, sheet_name=sheet_name, header=0, skiprows=[0, 1, 2, 3, 4, 6, 7, 8, 9]
)
df.rename(
    columns={
        "Handlungs- bedarf besteht  weitere Maßnahmen müssen umgesetzt werden \n(X)": "kein_handlungsbedarf"
    },
    inplace=True,
)
df.drop(df.index[df["kein_handlungsbedarf"] != "X"], inplace=True)

Row_list = df.to_numpy().tolist()

# print(Row_list)

while True:
    try:
        # Choose DEV or PROD
        value_env = input(
            "Please enter 'DEV' if you want to create measures on DEV or 'PROD' if you want to create measures on PROD.\n"
        )
        # Check that DEV or PROD has been chosen
        assert (value_env == "DEV") or (value_env == "PROD")
        break
    except AssertionError as wrong:
        logger.error("You must either select DEV or PROD. Please try again.")

customerID = input("Please enter the customer ID.\n")

# Check if customer is correct TODO: Handling Exceptions -> Customer does not exist

while True:

    if value_env == "PROD":
        CustomerURL = "{}/{}".format(PROD_URL, customerID)
    elif value_env == "DEV":
        CustomerURL = "{}/{}".format(DEV_URL, customerID)

    get_customer = requests.get(CustomerURL, data=customerID, headers=headers)
    cus_response = get_customer.json()
    print(
        "Is this the customer you want to upload measures for:",
        cus_response["data"]["name"],
    )
    check_customer_id = input(
        "Please type in 'Yes', if the customer is correct or 'No' if this is the wrong customer.\n"
    )

    if check_customer_id == "Yes":
        break
    elif check_customer_id == "No":
        customerID = input("Please enter the customer ID.\n")
    else:
        print("Incorrect response.")

if value_env == "PROD":
    myURL = "{}/{}/{}".format(PROD_URL, customerID, "measures")
elif value_env == "DEV":
    myURL = "{}/{}/{}".format(DEV_URL, customerID, "measures")

# Id of the division
division_id = input("Please enter the id of the division.\n")

# Check if division is correct TODO: Handling Exceptions -> Division does not exist

while True:

    if value_env == "PROD":
        DivisionURL = "{}/{}/{}/{}".format(
            PROD_URL, customerID, "facilities", division_id
        )
    elif value_env == "DEV":
        DivisionURL = "{}/{}/{}/{}".format(
            DEV_URL, customerID, "facilities", division_id
        )

    division_payload = {
        "facility_id": division_id,
        "customer_id": customerID,
    }

    get_division = requests.get(DivisionURL, data=division_payload, headers=headers)
    print("JSON Response ", get_division.json())
    div_response = get_division.json()
    print(
        "Is this the division you want to upload measures for:",
        div_response["data"]["location"],
        div_response["data"]["operational_area"],
    )
    check_div_id = input(
        "Please type in 'Yes', if the division is correct or 'No' if this is the wrong division.\n"
    )

    if check_div_id == "Yes":
        break
    elif check_div_id == "No":
        division_id = input("Please enter the id of the division.\n")
    else:
        print("Incorrect response.")

# Iterate
# TODO: Check if status, description and name have been filled out
for s in Row_list:

    """print("s0",s[0])
    print("s1",s[1])
    print("s2",s[2])
    print("s3",s[3])
    print("s4",s[4])
    print("s5",s[5])
    print("s6",s[6])
    print("s7",s[7])
    print("s8",s[8])
    print("s9",s[9])
    print("s10",s[10])
    print("s11",s[11])
    print("s12",s[12])"""

    if s[12] == "Erledigt":
        measure_status = True
    else:
        measure_status = False

    date = s[11]

    # Hazard must be a string. Then we can work with it.
    get_hazard = str(s[0])
    hazard = "".join(
        [a for a in get_hazard if not a.isdigit()]
    )  # Remove numbers like 1.1
    hazard = hazard.replace(".", "")  # Remove "."
    hazard = " ".join(hazard.split())  # Remove extra Spaces

    description = s[9]

    # Name must be a string. Then we can work with it.
    get_name = str([s[9]])
    # Remove everything after the first sentence.
    name = get_name.partition(".")[0]
    name = name.replace("[", "")  # Remove "["
    name = name.replace("'", "")  # Remove "'"
    name = " ".join(name.split())  # Remove extra Spaces

    pdf_status = False

    # Needed in order to assign the correct risk level
    remove_nan = "nan"

    if str(s[6]) not in remove_nan:
        risk_level = 1
    elif str(s[7]) not in remove_nan:
        risk_level = 2
    elif str(s[8]) not in remove_nan:
        risk_level = 3

    # Check that at least one risk level has been selected:
    if str(s[6]) == remove_nan and str(s[7]) == remove_nan and str(s[8]) == remove_nan:
        print("The risk level is missing for the following hazard:", s[0])
        risk_level = input("Please enter a risk level (1, 2 or 3)\n")
    # Check that only one risk level has been selected:
    if (
        (str(s[6]) != remove_nan and str(s[7]) != remove_nan)
        or (str(s[7]) != remove_nan and str(s[8]) != remove_nan)
        or (str(s[6]) != remove_nan and str(s[8]) != remove_nan)
    ):
        print("You have selected multiple risk levels for the following hazard:", s[0])
        risk_level = input("Please only choose one risk level(1, 2 or 3):\n")

    # Get the risk id
    get_first_char = str(s[0])
    first_char = get_first_char[0]

    if first_char == "" or first_char == "nan":
        print("The numbering seems to be broken / missing at:", s[0])
        first_char = input(
            "Please enter the correct number, e.g. such as in '1.1 ungeschützte bewegte Maschinenteile'):\n"
        )

    # DEV IDs
    if value_env == "DEV":

        if first_char == "1":
            # Mechanische G.
            risk_id = "49c06088-4a33-4694-9b26-6349e637fa40"
        elif first_char == "2":
            # Elektrische G.
            risk_id = "e278b2c7-47e0-4b40-ad83-7cf3bd076779"
        elif first_char == "3":
            # Gefahrstoffe
            risk_id = "2d4c9161-2d44-4aea-88e3-45249979ccbe"
        elif first_char == "4":
            # Gefahr und Biostoffe
            risk_id = "6946eb92-a7c7-4398-9f9a-dbe04a4809ae"
        elif first_char == "5":
            # Brandschutz
            risk_id = "e2def59b-0c0b-4dc4-8fb9-a2a122dbcf7c"
        elif first_char == "6":
            # Thermisch gibt es nicht, also Gefahrstoffe
            risk_id = "2d4c9161-2d44-4aea-88e3-45249979ccbe"
        elif first_char == "7":
            # Physikalisch gibt es nicht, also Allgemein
            risk_id = "e8212e55-ebad-4a6c-87d3-445696104e48"
        elif first_char == "8":
            # Arbeitsumgebung
            risk_id = "87760a21-4f0a-4da9-bf95-c176c75aff09"
        elif first_char == "9":
            # Physische Belastung gibt es nicht, also Arbeitsumgebung
            risk_id = "87760a21-4f0a-4da9-bf95-c176c75aff09"
        elif first_char == "10":
            # Psychisch
            risk_id = "a3a8725b-f1d3-4e76-ae13-da17f850dcb1"
        else:
            # Sonstiges, also Allgemein
            risk_id = "e8212e55-ebad-4a6c-87d3-445696104e48"
    else:
        risk_id = "fb058675-aec9-4a09-bcd7-d0adb256314e"  # Allgemein
        # TODO: PROD IDs und korrekte DEV IDs

    source = gbu_name

    # TODO: Check if Date is valid / if the format is correct
    format = "%d.%m.%Y"
    date = datetime.strptime(date, format)
    date = date.date()
    date = date.isoformat()

    body = {
        "done": measure_status,
        "due_date": date,
        "facility": {"id": division_id},
        "hazard": hazard,
        "hazard_description": description,
        "name": name,
        "pdf_sent": pdf_status,
        "risk_group": {"id": risk_id},
        "risk_level": risk_level,
        "source": source,
    }

    requestBody = json.dumps(body)

    # print(requestBody)

    # Do the request
    measure = requests.post(myURL, data=requestBody, headers=headers)
    print("Status Code", measure.status_code)
    print("JSON Response ", measure.json())
