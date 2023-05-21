#include <stdio.h>
#include <stdlib.h>
#include <conio.h>

void main()
{
    printf("Appuyez sur une touche pour continuer: \n");
    getch();
    system("python -m pip install --upgrade pip");
    system("python -m pip install pandas matplotlib numpy openpyxl");
    printf("tous est bien installes appuyez sur une touche pour continuer: \n");
    getch();
    printf("Vous voulez lancez l'application? Y/N  \n");
    char r;
    scanf("%c",&r);
    if (r == 'y' || r == 'Y') {
        system("app.exe");
    }
    else
    {
        printf("Appuyez sur une touche pour quitter: \n");
    }
}