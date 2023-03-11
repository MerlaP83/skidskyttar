import sqlite3, os, sys
from prettytable import PrettyTable

# Connect to the SQLite database
conn = sqlite3.connect('skidskytte.db')

# Define the columns to display in the table
columns = ['id', 'namn', 'nation', 'alder', 'kon', 'liggande_traffsakerhet', 'staende_traffsakerhet', 'tidstillagg', 'liggande_tid', 'staende_tid']

# Define the SQL query to get all skidskyttar data
sql = 'SELECT {} FROM skidskyttar'.format(', '.join(columns))

# Prompt the user to choose a gender option
gender_option = input('Visa alla skidskyttar (a), manliga skidskyttar (m) eller kvinnliga skidskyttar (f)? Gå tillbaka till huvudmenyn (h). Välj: ')

# Modify the SQL query based on the user's gender option
if gender_option == 'm':
    sql += ' WHERE kon = "m"'
elif gender_option == 'f':
    sql += ' WHERE kon = "f"'
elif gender_option == "h":
    os.system("index.py")

# Execute the query and fetch all results
cursor = conn.execute(sql)
skidskyttar = cursor.fetchall()

# Create a PrettyTable instance with the column names
table = PrettyTable(field_names=columns)

# Add each skidskytt to the table as a row
for skidskytt in skidskyttar:
    table.add_row(skidskytt)

# Print the table
print(table)

# Define the sorting options available to the user
sorting_options = {
    '1': ('namn', 'ASC'),
    '2': ('tidstillagg', 'ASC'),
    '3': ('liggande_traffsakerhet', 'DESC'),
    '4': ('staende_traffsakerhet', 'DESC'),
    '5': ('liggande_tid', 'ASC'),
    '6': ('staende_tid', 'ASC'),
    '7': ('alder', 'ASC'),
    '8': ('nation', 'ASC'),
}

# Define the starting and ending indices for displaying the skidskyttar
start = 0
end = 25

while True:
    # Create a sublist of skidskyttar to display based on the starting and ending indices
    skidskyttar_to_display = skidskyttar[start:end]

    # Create a PrettyTable instance with the column names
    table = PrettyTable(field_names=columns)

    # Add each skidskytt to the table as a row
    for skidskytt in skidskyttar_to_display:
        table.add_row(skidskytt)

    # Print the table
    print(table)

    # Prompt the user to choose a sorting option or go back to gender selection
    option = input('Sortera efter: \n1. Namn\n2. Tidstillägg\n3. Liggande träffsäkerhet\n4. Stående träffsäkerhet\n5. Liggande tid\n6. Stående tid\n7. Ålder\n8. Nation\nb. Gå tillbaka till könsvalet\ns. Visa nästa 25 åkare\nn. Visa föregående 25 åkare\nh. Huvudmenyn\nf. Avsluta programmet\nVälj: ')

    if option == 'b':
        # Go back to gender selection
        gender_option = input('Visa alla skidskyttar (a), manliga skidskyttar (m) eller kvinnliga skidskyttar (f)? Gå tillbaka till huvudmenyn (h). Välj: ')

        # Modify the SQL query based on the user's gender option
        if gender_option == 'm':
            sql = 'SELECT {} FROM skidskyttar WHERE kon = "m"'.format(', '.join(columns))
        elif gender_option == 'f':
            sql = 'SELECT {} FROM skidskyttar WHERE kon = "f"'.format(', '.join(columns))
        elif gender_option == "h":
            os.system("index.py")
        else:
            sql = 'SELECT {} FROM skidskyttar'.format(', '.join(columns))

        # Execute the query and fetch all results
        cursor = conn.execute(sql)
        skidskyttar = cursor.fetchall()

        # Reset the starting and ending indices
        start = 0
        end = 25

    elif option == 'q':
        # Exit the program
        print('Avslutar programmet.')
        break
    elif option == 'n':
        # Move to the next 25 skidskyttar
        if end < len(skidskyttar):
            start += 25
            end += 25
        else:
            print('Det finns inga fler åkare att visa.')
    elif option == 'f':
        # Move to the previous 25 skidskyttar
        if start >= 25:
            start -= 25
            end -= 25
        else:
            print('Det finns inga tidigare åkare att visa.')
    elif option == 'h':
        # Go back to main menu and run index.py
        print('Går tillbaka till huvudmenyn...')
        os.system('index.py')
        sys.exit()
    elif option not in sorting_options:
        # Invalid input
        print('Felaktig input. Försök igen.')
    else:
        # Sort the skidskyttar based on the user's sorting option
        column, order = sorting_options[option]
        skidskyttar.sort(key=lambda x: x[columns.index(column)], reverse=True if order == 'DESC' else False)

        # Reset the starting and ending indices
        start = 0
        end = 25