import os,sqlite3

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')


def main_menu():
    while True:
        clear_screen()
        print("SKIDSKYTTE SPEL")
        print("===============")
        print("1. Ny tävling")
        print("2. Se resultat")
        print("3. Se skidskyttar")
        print("4. Nollställ data")
        print("5. Avsluta")

        try:
            choice = int(input("Välj ett alternativ (1-5): "))
        except ValueError:
            print("Felaktig input. Försök igen.")
            input("Tryck på Enter för att fortsätta...")
            continue

        if choice == 1:
            os.system("sql-spel.py")
        elif choice == 2:
            os.system("resultat.py")
        elif choice == 3:
            os.system("skidskyttar.py")
        elif choice == 4:
            # Clear all entries from the resultat table and reset vpoang and starttid for all skidskyttar
            print("Rensar alla resultat och nollställer vpoang och starttid för alla skidskyttar.")
            confirm = input("Är du säker? (Ja/Nej): ")
            if confirm.lower() == "ja":
                # Connect to the database and execute the SQL statements to clear the resultat table and reset the vpoang and starttid columns in the skidskyttar table
                conn = sqlite3.connect("skidskytte.db")
                cursor = conn.cursor()
                cursor.execute("DELETE FROM resultat")
                cursor.execute("UPDATE skidskyttar SET vpoang = 0, starttid = 0")
                conn.commit()
                print("Alla resultat har raderats och vpoang och starttid har nollställts för alla skidskyttar.")
            else:
                print("Avbryter rensning av resultat och nollställning av vpoang och starttid.")
            input("Tryck på Enter för att fortsätta...")
        elif choice == 5:
            clear_screen()
            print("Tack för att du spelade! Hejdå!")
            break
        else:
            print("Ogiltigt val. Försök igen.")
            input("Tryck på Enter för att fortsätta...")


if __name__ == "__main__":
    main_menu()
