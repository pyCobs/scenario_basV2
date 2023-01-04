import datetime
from datetime import date


def addYears(d, years):
    try:
#Return same day of the current year
        return d.replace(year=d.year + years)
    except ValueError:
#If not same day, it will return other, i.e.  February 29 to March 1 etc.
        return d + (date(d.year + years, 1, 1) - date(d.year, 1, 1))


def decaler(list_valeur, decalage):
    nb_annees = 10
    list_valeur_decalees = []
    somme = 0

    if decalage >= 0:
        for annee in range(0, nb_annees):
            if annee < decalage:
                list_valeur_decalees.append(0)
                somme += int(list_valeur[nb_annees - annee - 1])
            else:
                list_valeur_decalees.append(list_valeur[annee - decalage])
        list_valeur_decalees[-1] = int(list_valeur_decalees[-1]) + somme

    else:  # décalages négatifs
        for annee in range(0, nb_annees):
            if annee > (nb_annees + decalage - 1):
                list_valeur_decalees.append(0)
            else:
                list_valeur_decalees.append(str(list_valeur[annee - decalage]))

    list_valeur_decalees = [str(x) for x in list_valeur_decalees]

    return list_valeur_decalees


if __name__ == "__main__":
    list_valeur = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
    decalage = 3

    list_valeur_decalees = decaler(list_valeur, decalage)
    print(list_valeur_decalees)