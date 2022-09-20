from asyncio.log import logger
from tkinter import Entry
import requests
import json
from requests.structures import CaseInsensitiveDict
import pandas as pd
from datetime import datetime
import locale
import openpyxl  # import needed for windows exe - otherwise the converter will miss it even though it should work via pandas

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
    file_name = input(
        "\nPlease enter the name of the Excel file. Don't forget the filename extension.\n"
    )
    try:
        open(file_name, "r")
        break
    except FileNotFoundError as e:
        logger.error("\nThe file does not exist. Please try again.")


while True:

    # Get the name of the sheet and check if it exists
    sheet_name = input(
        "\nPlease enter the name of the sheet (tab) that contains the risk assessment.\n"
    )
    try:
        get_first_row = pd.read_excel(
            file_name, sheet_name=sheet_name, nrows=1, usecols="A"
        )
        break
    except ValueError as error:
        logger.error("\nThe sheet does not exist. Please try again.")

# Open the file and then get the first row (this is needed for the source field)
get_first_row = pd.read_excel(file_name, sheet_name=sheet_name, nrows=1, usecols="A")
get_gbu_name = str(get_first_row)

# Get only the name of the risk assessment
gbu_name = get_gbu_name.partition("\n")[0]
gbu_name = " ".join(gbu_name.split())  # Remove extra Spaces

# Filter rows of the Excel sheet to extract the necessary information
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
            "\nPlease enter 'DEV' if you want to create measures on DEV or 'PROD' if you want to create measures on PROD.\n"
        )
        # Check that DEV or PROD has been chosen
        assert (value_env.upper() == "DEV") or (value_env.upper() == "PROD")
        break
    except AssertionError as wrong:
        logger.error("\nYou must either select DEV or PROD. Please try again.")

# Check if customer exists and if the customer is correct
customerID = input("\nPlease enter the customer ID.\n")

while True:

    if value_env.upper() == "PROD":
        CustomerURL = "{}/{}".format(PROD_URL, customerID)
    elif value_env.upper() == "DEV":
        CustomerURL = "{}/{}".format(DEV_URL, customerID)

    get_customer = requests.get(CustomerURL, data=customerID, headers=headers)
    cus_response = get_customer.json()

    try:
        print(
            "\nIs this the customer you want to upload measures for:",
            cus_response["data"]["name"],
        )
    except KeyError as keyerror:
        raise Exception(
            "\nThe customer does not exist",
        ) from None

    check_customer_id = input(
        "\nPlease type in 'Yes', if the customer is correct or 'No' if this is the wrong customer.\n"
    )

    if check_customer_id.lower() == "yes":
        break
    elif check_customer_id.lower() == "no":
        customerID = input("\nPlease enter the customer ID.\n")
    else:
        logger.error("\nInvalid response.")

if value_env.upper() == "PROD":
    myURL = "{}/{}/{}".format(PROD_URL, customerID, "measures")
elif value_env.upper() == "DEV":
    myURL = "{}/{}/{}".format(DEV_URL, customerID, "measures")

# Id of the division
division_id = input("\nPlease enter the id of the division.\n")

# Check if division exists and if it is correct
while True:

    if value_env.upper() == "PROD":
        DivisionURL = "{}/{}/{}/{}".format(
            PROD_URL, customerID, "facilities", division_id
        )
    elif value_env.upper() == "DEV":
        DivisionURL = "{}/{}/{}/{}".format(
            DEV_URL, customerID, "facilities", division_id
        )

    division_payload = {
        "facility_id": division_id,
        "customer_id": customerID,
    }

    try:
        get_division = requests.get(DivisionURL, data=division_payload, headers=headers)
        div_response = get_division.json()
        print(
            "\nIs this the division you want to upload measures for:",
            div_response["data"]["location"],
            div_response["data"]["operational_area"],
        )
    except requests.JSONDecodeError as json_decode_error:
        raise Exception(
            "\nThe division does not exist",
        ) from None

    check_div_id = input(
        "\nPlease type in 'Yes', if the division is correct or 'No' if this is the wrong division.\n"
    )

    if check_div_id.lower() == "yes":
        break
    elif check_div_id.lower() == "no":
        division_id = input("\nPlease enter the id of the division.\n")
    else:
        logger.error("\nInvalid response.")

# Get all existing measures for the customer and the division

if value_env.upper() == "PROD":
    MeasuresURL = "{}/{}/{}".format(PROD_URL, customerID, "measures?")
elif value_env.upper() == "DEV":
    MeasuresURL = "{}/{}/{}".format(DEV_URL, customerID, "measures?")

measure_payload = {
    "locations": div_response["data"]["location"],
    "operational_areas": div_response["data"]["operational_area"],
}

get_all_measures = requests.get(MeasuresURL, measure_payload, headers=headers)
all_measures = get_all_measures.json()

measure_list = []

for entry in all_measures["data"]:
    measure_output = entry["name"]
    measure_list.append(measure_output)

# Iterate
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

    measure_status = s[12]

    while True:

        if str(measure_status) == "Erledigt" or str(measure_status) == "erledigt":
            measure_status = True
            break
        elif str(measure_status) == "Offen" or str(measure_status) == "offen":
            measure_status = False
            break
        else:
            print(
                "\nThe status of the following hazard is either missing or not valid:",
                s[0],
            )
            measure_status = input(
                "\nPlease type in the status ('Erledigt' or 'Offen'):\n"
            )

    # Hazard must be a string. Then we can work with it.
    get_hazard = str(s[0])
    hazard = "".join(
        [a for a in get_hazard if not a.isdigit()]
    )  # Remove numbers like 1.1
    hazard = hazard.replace(".", "")  # Remove "."
    hazard = " ".join(hazard.split())  # Remove extra Spaces

    # Needed for some comparisons
    empty_nan = "nan"
    description = str(s[9])

    if description == empty_nan:

        print("\nThe description is missing for the following hazard:", s[0])

        while True:
            description = input("\nPlease type in a description:\n")
            print("\nIs the description correct?\n", description)
            check_des = input(
                "\nPlease type in 'Yes', if the description is correct or 'No' if you want to change it.\n"
            )
            if check_des.lower() == "yes":
                if description == empty_nan:
                    print("\nThe description is still missing.\n")
                else:
                    break

    # Name must be a string. Then we can work with it.
    get_name = description

    if get_name == empty_nan:

        print("\nThe name is missing for the following hazard:", s[0])

        while True:
            get_name = input("Please type in a name:\n")
            print("\nIs the name correct?\n", get_name)
            check_name = input(
                "\nPlease type in 'Yes', if the name is correct or 'No' if you want to change it.\n"
            )
            if check_name.lower() == "yes":
                if name == empty_nan:
                    print("\nThe name is still missing.\n")
                else:
                    break

    # Remove everything after the first sentence.
    name = get_name.partition(".")[0]
    name = get_name.partition("\n")[0]  # Let's try out if this makes sense.
    name = name.replace("[", "")  # Remove "["
    name = name.replace("]", "")  # Remove "["
    name = name.replace("'", "")  # Remove "'"
    name = " ".join(name.split())  # Remove extra Spaces

    # Asks to skip hazard if it already exists
    if name in measure_list:
        print(
            "\nA measure with the following name: ",
            "'",
            name,
            "'",
            " is already in the portal for this location and operational area.",
        )
        create_or_not = input(
            "\nDo you still want to create the measure? Please type in 'Yes' or 'No'.\n"
        )

        if create_or_not.lower() == "no":
            continue
        elif (create_or_not.lower() != "no") and (create_or_not.lower() != "yes"):
            logger.error("\nInvalid response.")

    pdf_status = False

    if str(s[6]) not in empty_nan:
        risk_level = 1
    elif str(s[7]) not in empty_nan:
        risk_level = 2
    elif str(s[8]) not in empty_nan:
        risk_level = 3

    # Check that at least one risk level has been selected:
    if str(s[6]) == empty_nan and str(s[7]) == empty_nan and str(s[8]) == empty_nan:

        print("\nThe risk level is missing for the following hazard:", s[0])

        while True:
            risk_level = int(input("Please enter a risk level (1, 2 or 3)\n"))
            check_risk_level = input(
                "\nPlease type in 'Yes', if the risk level is correct or 'No' if you want to change it.\n"
            )
            if check_risk_level.lower() == "yes":

                if risk_level == empty_nan:
                    print("\nThe risk level is still missing.")
                elif ((risk_level != 1) and (risk_level != 2) and (risk_level != 3)):
                    print("\nThe risk level is invalid.")
                else:
                    break

    # Check that only one risk level has been selected:
    if (
        (str(s[6]) != empty_nan and str(s[7]) != empty_nan)
        or (str(s[7]) != empty_nan and str(s[8]) != empty_nan)
        or (str(s[6]) != empty_nan and str(s[8]) != empty_nan)
    ):
        print(
            "\nYou have selected multiple risk levels for the following hazard:", s[0]
        )
        while True:
            risk_level = int(input("Please enter a risk level (1, 2 or 3):\n"))
            check_risk_level = input(
                "\nPlease type in 'Yes', if the risk level is correct or 'No' if you want to change it.\n"
            )
            if check_risk_level.lower() == "yes":

                if risk_level == empty_nan:
                    print("\nThe risk level is still missing.")
                elif ((risk_level != 1) and (risk_level != 2) and (risk_level != 3)):
                    print("\nThe risk level is invalid.")
                else:
                    break

    # Get the risk id
    get_first_char = str(s[0])
    first_char = get_first_char[0]

    # Check number
    if (
        (first_char == "")
        or (first_char == "nan")
        or (str(first_char)).isnumeric() == False
    ):
        print(
            "\nThe numbering seems to be missing at:",
            s[0],
            ". Without the numbering we cannot select the correct risk id.",
        )

        while True:
            first_char = input(
                "\nPlease enter the first character, e.g. for '1.1 ungeschützte bewegte Maschinenteile' you would enter '1'.):\n"
            )
            check_first_char = input(
                "\nPlease type in 'Yes', if the risk level is correct or 'No' if you want to change it.\n"
            )
            if check_first_char.lower() == "yes":

                if (
                    (first_char == "")
                    or (first_char == "nan")
                    or (str(first_char)).isnumeric() == False
                ):
                    print("\nThe risk level is still broken.")
                else:
                    break

    # DEV IDs
    if value_env.upper() == "DEV":

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
            # Thermische G.
            risk_id = "4fbbb167-95a2-4b81-92ca-4530c29bc957"
        elif first_char == "7":
            # Physikalische Einwirkungen
            risk_id = "c9da8c0e-a37c-46c4-be27-188dea32d25a"
        elif first_char == "8":
            # Arbeitsumgebungsbedingungen
            risk_id = "c3feb662-75e2-4548-b0eb-f1d697416767"
        elif first_char == "9":
            # Physische Belastung
            risk_id = "e7d2edf7-7591-4fb3-b60f-7c1f7994a958"
        elif first_char == "10":
            # Psychische Faktoren
            risk_id = "a3a8725b-f1d3-4e76-ae13-da17f850dcb1"
        elif first_char == "11":
            # Sonstige Gefährdungen
            risk_id = "a506f7b8-336d-40b4-aa0a-e0c4c6ba8de6"
        else:
            # Nicht valide
            print("Invalid risk id.")
            raise

    elif value_env.upper() == "PROD":

        if first_char == "1":
            # Mechanische G.
            risk_id = "3a6d0f2c-00aa-47a8-85a0-6c73a28181b9"
        elif first_char == "2":
            # Elektrische G.
            risk_id = "951ef441-fd9d-45dc-bf8b-343e5019e2db"
        elif first_char == "3":
            # Gefahrstoffe
            risk_id = "2421bcdb-322b-4b2b-b868-32fc05c07ce7"
        elif first_char == "4":
            # Gefahr und Biostoffe
            risk_id = "b8a2f7a5-325f-4c6b-8727-48dcc3406f9c"
        elif first_char == "5":
            # Brandschutz
            risk_id = "881ef746-efc0-40d6-b67f-04846d89a35c"
        elif first_char == "6":
            # Thermische G.
            risk_id = "269cd96b-3d0c-4337-80d6-c97898e80890"
        elif first_char == "7":
            # Physikalische Einwirkungen
            risk_id = "53d0a16f-45d3-4073-896c-0263d8c552b1"
        elif first_char == "8":
            # Arbeitsumgebungsbedingungen
            risk_id = "52a1f928-0eb0-4fbf-a63f-2b975158bf0b"
        elif first_char == "9":
            # Physische Belastung
            risk_id = "f7167ec6-7cbe-4771-bc91-b7517ad19f55"
        elif first_char == "10":
            # Psychische Faktoren
            risk_id = "832b98b5-7a2e-4eb7-9fbd-881643157e16"
        elif first_char == "11":
            # Sonstige Gefährdungen
            risk_id = "09259b68-df01-4820-b6be-7441c7435c67"
        else:
            # Nicht valide
            print("Invalid risk id.")
            raise

    source = gbu_name

    date = s[11]
    format = "%d.%m.%Y"

    if str(date) != empty_nan:

        # Check for correct date format
        try:
            date = datetime.strptime(date, format)
            date = date.date()
            date = date.isoformat()
        except ValueError as valueerror:
            raise Exception(
                "\nIncorrect data format. The date format should be 'DD.MM.YYYY' for the following hazard:",
                s[0],
            ) from None
        except TypeError as typeerror:
            raise Exception(
                "\nIncorrect data format. The date format should be 'DD.MM.YYYY' for the following hazard:",
                s[0],
            ) from None

    if str(date) == empty_nan:
        body = {
            "done": measure_status,
            "facility": {"id": division_id},
            "hazard": hazard,
            "hazard_description": description,
            "name": name,
            "pdf_sent": pdf_status,
            "risk_group": {"id": risk_id},
            "risk_level": risk_level,
            "source": source,
        }

    else:
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
    if str(measure.status_code) == "201":
        print("\nThe following measure has been created:", name)
    else:
        print("\nThe following measure could not be created:", name)
        print("Status Code", measure.status_code)
        print("JSON Response ", measure.json())

""" print("Status Code", measure.status_code)
    print("JSON Response ", measure.json()) """

input("\nPress any key to exit the script...")
