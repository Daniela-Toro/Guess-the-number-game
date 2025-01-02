# IMPORTAR FUNCIONES
from colorama import Fore, Style
from random import randint
from datetime import datetime
import openpyxl
from openpyxl import Workbook
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import pygame
import time

#FUNCION PARA MOSTRAR EL MENU PRINCIPAL
def Menu():
    print(Fore.BLUE + "▓" * 50)
    print(Fore.BLUE + Style.BRIGHT + "¡QUE COMIENCE EL JUEGO: ADIVINA EL NÚMERO! \U0001F40D \U0001F9D1\U0000200D\U0001F4BB")
    print(Fore.BLUE + "▓" * 50)
    print(Fore.CYAN + "MENÚ:".center(40))
    print(
        Fore.CYAN + "\t1. \U0001F3AE PARTIDA MODO SOLITARIO\n" +
        "\t2. \U0001F465 PARTIDA 2 JUGADORES\n" +
        "\t3. \U0001F4CA ESTADÍSTICAS\n" +
        "\t4. \U0001F6AA SALIR" +
        Style.RESET_ALL
    )

# FUNCION PARA REPRODUCIR SONIDO DE GANAR O PERDER
def animation_game (url_photo, url_sound):
# Mostrar Sonido
    pygame.init()
    pygame.mixer.init()
    sound = pygame.mixer.Sound(url_sound)
    sound.play()
    time.sleep(2)
    pygame.quit()
# Mostrar Imagen
    img = mpimg.imread(url_photo)
    plt.imshow(img)
    plt.axis('off')
    plt.show()

# FUNCION PARA VALIDAR ENTRADAS
def validation(value, min, max):
    while value < min or value > max:
        value = int(input(Fore.BLUE + f"UPS! TIENES QUE SELECCIONAR UNA OPCIÓN ENTRE {min} Y {max}: "))
    return value

# FUNCION PARA MOSTRAR SUBMENU Y ELEGIR DIFICULTAD
def submenu():
    print(Fore.LIGHTCYAN_EX + "SELECCIONA EL NIVEL DE DIFICULTAD".center(40))
    print(
        Fore.LIGHTCYAN_EX + "\t1. \U0001F480 FÁCIL\n" +
        "\t2. \U0001F480 \U0001F480 MEDIO\n" +
        "\t3. \U0001F480 \U0001F480 \U0001F480 DIFÍCIL\n" +
        Style.RESET_ALL
    )
    difficulty = int(input(Fore.BLUE + "SELECCIONA EL NIVEL DE DIFICULTAD: "))
    difficulty = validation(difficulty,1,3)
    return {1: 20, 2: 12, 3: 5}[difficulty]

# FUNCION LOGICA DEL JUEGO
def play_game(max_attempts,unknown_number,name):
    print(Fore.LIGHTWHITE_EX + f"{name}, TIENES {max_attempts} INTENTOS PARA ADIVINAR!.")
    attempts = 0
    guess_number = 0
    win = False
    while (max_attempts > attempts) and (guess_number != unknown_number):
        guess_number = int(input(Fore.BLUE + "INTRODUCE UN NÚMERO: "))
        guess_number = validation(guess_number, 1, 1000)
        if guess_number == unknown_number:
            animation_game('win.jpg','win.wav')
            print(f"FELICIDADES {name}! HAS GANADO!")
            win = True
        elif guess_number > unknown_number:
            print("EL NÚMERO A ADIVINAR ES MENOR")
        else:
            print("EL NÚMERO A ADIVINAR ES MAYOR")
        attempts += 1
        if win == False:
            print(f'TE QUEDAN {max_attempts - attempts} INTENTOS DISPONIBLES')
    if max_attempts == attempts:
        print(Fore.RED + f"{name} HAS SUPERADO EL NÚMERO MÁXIMO DE INTENTOS PERMITIDOS\nSUERTE PARA LA PRÓXIMA")
        animation_game('lose.jpg','game_over.wav')
    save_statics(name, attempts, win, unknown_number, max_attempts)

# FUNCION PARA MODO SOLITARIO
def one_player():
    print(Fore.BLUE + "=" * 45)
    print(Fore.LIGHTWHITE_EX + "\U0001F3AE PARTIDA MODO SOLITARIO.".center(40))
    print(Fore.LIGHTWHITE_EX + f"DESCUBRE EL NÚMERO QUE HEMOS ESCONDIDO ENTRE 1 Y 1000.")
    max_attempts = submenu()
    unknown_number = randint(1, 1000)
    print(unknown_number)
    name = input("INTRODUCE TU NOMBRE O ALIAS: ").lower()
    play_game(max_attempts, unknown_number,name)

# FUNCION PARA MODO 2 JUGADORES
def two_players():
    print(Fore.BLUE + "=" * 45)
    print(Fore.LIGHTWHITE_EX + "\U0001F465 PARTIDA 2 JUGADORES.".center(40))
    max_attempts = submenu()
    print(Fore.LIGHTWHITE_EX + "JUGADOR 1: ADIVINARÁS UN NÚMERO ENTRE 1 Y 1000 QUE EL JUGADOR 2 SELECCIONARÁ ")
    name = input("JUGADOR 1, INTRODUCE TU NOMBRE O ALIAS: ").lower()
    name2 = input("JUGADOR 2, INTRODUCE TU NOMBRE O ALIAS: ").lower()
    unknown_number = validation(int(input(Fore.LIGHTWHITE_EX + f"{name2}: INSERTA UN NÚMERO ENTRE 1 Y 1000: ")), 1, 1000)
    play_game(max_attempts, unknown_number,name)
# FUNCION PARA GUARDAR ESTADISTICAS
def save_statics(name,attempts,win,unknown_number,max_attempts):
    try:
        wb = openpyxl.load_workbook("GAME_STATICS.xlsx")
        sheet = wb.active
    except FileNotFoundError:
        wb = Workbook()
        sheet = wb.active
        sheet.title = "statics"
        sheet.append(["NOMBRE", "GANADOR", "NÚMERO_SECRETO", "INTENTOS_UTILIZADOS", "INTENTOS_TOTALES", "FECHA"])
        print(Fore.RED + "ARCHIVO GAME_STATICS.xlsx CREADO")
    dt = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    sheet.append([name, win, unknown_number, attempts, max_attempts, dt])
    wb.save("GAME_STATICS.xlsx")
    print(Fore.RED + "RESULTADO DE LA PARTIDA GUARDADO EN ESTADÍSTICAS")

#FUNCION PARA MOSTRAR ESTADISTICAS
def show_statics ():
    print(Fore.BLUE + "=" * 45)
    print(Fore.LIGHTWHITE_EX + "\U0001F4CA ESTADÍSTICAS.".center(40))
    try:
        wb = openpyxl.load_workbook("GAME_STATICS.xlsx")
        Hoja = wb['statics']
        print(Fore.LIGHTCYAN_EX + "MENÚ".center(40))
        print(Fore.RED + "NOTA: Las partidas aparecen guardadas en el archivo GAME_STATICS.xlsx")
        print(Fore.LIGHTCYAN_EX + "\t1. ESTADÍSTICAS GENERALES\n" + "\t2. ESTADÍSTICAS POR USUARIO\n" + Style.RESET_ALL)
        option = int(input(Fore.BLUE + "SELECCIONE QUÉ DESEA HACER: "))
        option = validation(option, 1, 2)
        if option == 1:
            print(Fore.LIGHTCYAN_EX + "\nESTADÍSTICAS GENERALES".center(45) + Style.RESET_ALL)
            player_wins = {}
            for cell in Hoja[1]:
                print(cell.value, end=" ")
            print()
            for row in Hoja.iter_rows(min_row=2, values_only=True):
                name, win, unknown_number, attempts, max_attempts, date = row
                for cell in row:
                    print(cell, end=" ")
                print()
                if not name in player_wins:
                    player_wins[name] = 0
                if win == True:
                    player_wins[name] += 1
            names = list(player_wins.keys())
            wins = list(player_wins.values())
            plt.figure(figsize=(10, 6))
            plt.bar(names, wins, color='green')
            plt.xlabel('Nombres de Jugadores')
            plt.ylabel('Partidas Ganadas')
            plt.title('Partidas Ganadas por Jugador')
            plt.xticks(rotation=45, ha="right")
            plt.tight_layout()
            plt.show()
        else:
            total_wins = 0
            total_losses = 0
            user_exists = False
            user = input('INTRODUZCA EL NOMBRE DEL USUARIO QUE DESEA BUSCAR: ')
            print(Fore.GREEN + f"\nInformación del usuario {user}:" + Style.RESET_ALL)
            for row in Hoja.iter_rows(min_row=2, values_only=True):
                name, win, unknown_number, attempts, max_attempts, date = row
                if name.lower() == user.lower():
                    user_exists = True
                    for cell in row:
                        print(cell, end=" ")
                    print()
                    if win == True:
                        total_wins += 1
                    else:
                        total_losses += 1
            if user_exists == True:
                categories = ['Partidas Ganadas', 'Partidas Perdidas']
                values = [total_wins, total_losses]
                plt.figure(figsize=(8, 6))
                plt.bar(categories, values, color=['green', 'red'])
                plt.xlabel('Categorías')
                plt.ylabel('Cantidad de Partidas')
                plt.title(f'Estadísticas de {user}')
                plt.show()
            else:
                print(f'EL USUARIO {user} NO EXISTE')
    except FileNotFoundError:
        print("ERROR: ARCHIVO NO ENCONTRADO. ASEGÚRATE DE JUGAR AL MENOS UNA PARTIDA ANTES")

# PROGRAMA PRINCIPAL
selection = "0"
while selection != "4":
    Menu()
    selection = input(Fore.BLUE + "SELECCIONA UNA DE LAS OPCIONES ANTERIORES: ")
    if selection == "1":
        one_player()
    elif selection == "2":
        two_players()
    elif selection == "3":
        show_statics()
    elif selection == "4":
        print(Fore.RED + "FIN DEL JUEGO\nGRACIAS POR JUGAR, HASTA LA PRÓXIMA! \U0001F44B")
    else:
        print(Fore.RED + "=" * 45)
        print(Fore.RED + "\u274C ERROR: VALOR INSERTADO NO VÁLIDO \u274C\n DEBE SELECCIONAR UNA OPCIÓN ENTRE 1 Y 4")
        print(Fore.RED + "=" * 45)