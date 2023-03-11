import sqlite3,os
from prettytable import PrettyTable

conn = sqlite3.connect("skidskytte.db")
cursor = conn.cursor()

# Define the columns to display in the result table
columns = [
    "Placering",
    "Tävling",
    "Total tid",
    "Tid sedan ledare",
    "Straffrundor",
]

def get_tavlingar():
    """Returns a list of unique tavling names from the database that have at least one row in the resultat table."""
    cursor.execute(
        "SELECT id, tavling, GROUP_CONCAT(namn, ', ') AS topp_3 "
        "FROM resultat "
        "WHERE placering <= 3 "
        "GROUP BY id"
    )
    return cursor.fetchall()

def display_resultat(tavling_id, tavling_namn, topp_3):
    """Displays the resultat table for the specified tavling."""
    cursor.execute(
        "SELECT placering, namn, total_tid, tid_sedan_ledare, straffrundor_totalt "
        "FROM resultat WHERE id = ? "
        "GROUP BY id, placering "
        "ORDER BY placering",
        (tavling_id,)
    )
    result = cursor.fetchall()

    # Create a PrettyTable instance with the column names
    table = PrettyTable(field_names=columns)

    # Add each resultat to the table as a row
    for row in result:
        total_tid = f"{int(row[2] // 60)}:{int(row[2] % 60):02d}"
        tid_sedan_ledare = row[3] if row[3] is not None else ""
        table.add_row([row[0], row[1], total_tid, tid_sedan_ledare, row[4]])

    # Print the table
    print(table)

def get_skidskyttar():
    """Returns a list of unique skidskyttar names from the database that have at least one row in the resultat table."""
    cursor.execute("SELECT id, namn FROM skidskyttar order by namn ASC")
    return cursor.fetchall()

def display_skidskytt_resultat(skidskytt_id, skidskytt_namn):
    """Displays the resultat table for the specified skidskytt."""
    while True:
        cursor.execute(
            "SELECT id, tavling, gren, placering, straffrundor_totalt, total_tid "
            "FROM resultat WHERE akarid = ? "
            "GROUP BY id",
            (skidskytt_id,)
        )

        result = cursor.fetchall()

        # Print the header
        print(f"Resultat för skidskytt {skidskytt_namn} (ID: {skidskytt_id}):")

        # Display a menu of available lopp
        print("Lopp:")
        for i, lopp in enumerate(result):
            total_tid = f"{int(lopp[5] // 60)}:{int(lopp[5] % 60):02d}"
            print(f"{i+1}. {lopp[1]}, {lopp[2].lower()}: #{lopp[3]} - {total_tid}({lopp[4]}).")

        # Prompt the user to choose a lopp or go back to the main menu
        while True:
            try:
                choice = int(input("Välj ett lopp att visa resultat för (eller 0 för att gå tillbaka till huvudmenyn): "))
                if choice < 0 or choice > len(result):
                    raise ValueError
                break
            except ValueError:
                print("Felaktig input. Försök igen.")

        if choice == 0:
            # Go back to the main menu
            return
        else:
            # Get the chosen lopp ID and display the corresponding resultat table
            lopp_id = result[choice - 1][0]
            cursor.execute(
                "SELECT placering, namn, total_tid, tid_sedan_ledare, straffrundor_totalt "
                "FROM resultat WHERE id = ? "
                "ORDER BY placering",
                (lopp_id,)
            )

            result = cursor.fetchall()

            # Create a PrettyTable instance with the column names
            columns = [
                "Placering",
                "Namn",
                "Total tid",
                "Tid sedan ledare",
                "Straffrundor",
            ]
            table = PrettyTable(field_names=columns)

            # Add each skidskytt to the table as a row
            for row in result:
                total_tid = f"{int(row[2] // 60)}:{int(row[2] % 60):02d}" if row[2] is not None else ""
                tid_sedan_ledare = row[3] if row[3] is not None else ""
                table.add_row([row[0], row[1], total_tid, tid_sedan_ledare, row[4]])

        # Print the table
        print(table)


while True:
    # Display a menu of available options
    print("Välj ett alternativ:")
    print("1. Visa resultat för en skidskytt")
    print("2. Visa resultat för en tävling")
    print("3. Avsluta programmet")

    # Prompt the user to choose an option
    while True:
        try:
            choice = int(input("Ange ditt val: "))
            if choice < 1 or choice > 3:
                raise ValueError
            break
        except ValueError:
            print("Felaktig input. Försök igen.")

    # Perform the chosen action
    if choice == 1:
        # Get the chosen skidskytt name and display the corresponding resultat table
        skidskyttar = get_skidskyttar()
        print("Skidskyttar:")
        for i, skidskytt in enumerate(skidskyttar):
            print(f"{i+1}. {skidskytt[1]} ({skidskytt[0]})")

        while True:
            try:
                choice = int(input("Välj en skidskytt att visa resultat för: "))
                if choice < 1 or choice > len(skidskyttar):
                    raise ValueError
                break
            except ValueError:
                print("Felaktig input. Försök igen.")

        skidskytt_id, skidskytt_namn = skidskyttar[choice - 1]
        display_skidskytt_resultat(skidskytt_id, skidskytt_namn)

    elif choice == 2:
        # Get the chosen tavling name and display the corresponding resultat table
        tavlingar = get_tavlingar()
        print("Tävlingar:")
        for i, tavling in enumerate(tavlingar):
            print(f"{i+1}. {tavling[1]} ({tavling[0]})")

        while True:
            try:
                choice = int(input("Välj en tävling att visa resultat för: "))
                if choice < 1 or choice > len(tavlingar):
                    raise ValueError
                break
            except ValueError:
                print("Felaktig input. Försök igen.")

        # Get the chosen tavling name and display the corresponding resultat table
        tavling_id, tavling_namn, topp_3 = tavlingar[choice - 1]
        print(f"\nResultat för tävling {tavling_namn} (ID: {tavling_id}):")
        display_resultat(tavling_id, tavling_namn, topp_3)

    elif choice == 3:
        # Exit the program
        print("Huvudmenyn")
        os.system("index.py")
        break

    else:
        print("Felaktig input. Försök igen.")
