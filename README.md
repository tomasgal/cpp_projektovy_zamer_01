Excel VBA Projekt: Dynamické Formátovanie a Vstupné Masky s VBA

Tento projekt demonštruje rôzne možnosti použitia VBA v Exceli na úpravu formátovania, interaktívne preklikávanie textu a prácu s dátumovými rozsahmi. Hlavné funkcie zahŕňajú formátovanie prepojené s VBA kódom, vytváranie vstupnej masky pre dátumové rozsahy a dynamické striedanie preškrtnutého textu.


Funkcie

    Formátovanie dátumového rozsahu v TextBoxe: Textové pole s VBA maskou, ktoré automaticky upravuje zobrazenie dátumov vo formáte "dd. mm. yyyy - dd. mm. yyyy".
    Dynamické preklikávanie preškrtnutého textu: Vytvára dvojice buniek, v ktorých môže byť preškrtnutý text len v jednej bunke z dvojice. Preklikaním bunky sa stav preškrtnutia strieda.
    Automatické formátovanie prázdnych buniek: Nastavuje text ako "Dátum: " v prípade, že bunka neobsahuje žiadne hodnoty.

Inštalácia

    Stiahnite alebo naklonujte tento projekt z GitHubu.
    Otvorte Excel súbor a povoľte makrá:
        Pri otvorení súboru kliknite na Povoliť obsah, aby ste mohli spustiť VBA skripty.
    Pre plnú funkcionalitu zapnite kartu Vývojár v Exceli:
        Prejdite na Súbor > Možnosti > Prispôsobiť pás s nástrojmi a začiarknite Vývojár.

Použitie
1. Formátovanie Dátumového Rozsahu v TextBoxe

    V textovom poli (TextBox ActiveX) môžete zadať dátumový rozsah vo formáte dd.mm.yyyy-dd.mm.yyyy.
    Po opustení poľa sa vstup automaticky preformátuje na "dd. mm. yyyy - dd. mm. yyyy".

2. Preklikávanie Preškrtnutého Textu

    Kliknutím na konkrétnu bunku v dvojiciach, ako napríklad A1 a B1, sa preškrtnutie textu prepína medzi týmito dvoma bunkami.
    Len jedna bunka z každej dvojice môže byť preškrtnutá naraz, čím sa zabezpečí konzistentné zobrazenie.

3. Automatické zobrazenie textu v prázdnych bunkách

    V bunkách nastavených pre dátum sa automaticky zobrazí text "Dátum: ", ak bunka neobsahuje žiadnu hodnotu.
    V prípade zadania dátumu sa text dynamicky spojí s hodnotou dátumu.

Podrobnosti VBA Skriptov
Formátovanie Dátumového Rozsahu v TextBoxe

Tento VBA kód spracováva formátovanie pre TextBox ActiveX s automatickým formátovaním po zadaní dátumu.

vba

Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim dateRange As String
    dateRange = Trim(TextBox1.Value)
    
    ' Skontrolujeme, či text zodpovedá formátu ##.##.####-##.##.####
    If dateRange Like "##.##.####-##.##.####" Then
        ' Preformátujeme na ##. ##. #### - ##. ##. ####
        TextBox1.Value = Left(dateRange, 2) & ". " & Mid(dateRange, 4, 2) & ". " & Mid(dateRange, 7, 4) & " - " & _
                         Mid(dateRange, 12, 2) & ". " & Mid(dateRange, 15, 2) & ". " & Right(dateRange, 4)
    Else
        MsgBox "Zadajte dátumové rozpätie vo formáte ##.##.####-##.##.#### bez medzier", vbExclamation
        Cancel = True ' Zostaneme v poli, ak formát nesúhlasí
    End If
End Sub

Interaktívne Preklikávanie Preškrtnutého Textu

Tento VBA kód zabezpečuje, že pre každú dvojicu buniek, ako napríklad A1 a B1, sa preškrtnutie textu strieda len medzi týmito dvoma bunkami.

vba

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Application.EnableEvents = False ' Vypneme udalosti, aby sme zabránili opakovaniu
    
    ' Definujeme dvojice buniek a prepíname preškrtnutie pre každú dvojicu
    If Not Intersect(Target, Me.Range("A1:B1")) Is Nothing Then
        Me.Range("A1").Font.Strikethrough = Not Me.Range("A1").Font.Strikethrough
        Me.Range("B1").Font.Strikethrough = Not Me.Range("A1").Font.Strikethrough
    ElseIf Not Intersect(Target, Me.Range("A2:B2")) Is Nothing Then
        Me.Range("A2").Font.Strikethrough = Not Me.Range("A2").Font.Strikethrough
        Me.Range("B2").Font.Strikethrough = Not Me.Range("A2").Font.Strikethrough
    ElseIf Not Intersect(Target, Me.Range("A3:B3")) Is Nothing Then
        Me.Range("A3").Font.Strikethrough = Not Me.Range("A3").Font.Strikethrough
        Me.Range("B3").Font.Strikethrough = Not Me.Range("A3").Font.Strikethrough
    ElseIf Not Intersect(Target, Me.Range("A4:B4")) Is Nothing Then
        Me.Range("A4").Font.Strikethrough = Not Me.Range("A4").Font.Strikethrough
        Me.Range("B4").Font.Strikethrough = Not Me.Range("A4").Font.Strikethrough
    End If
    
    Application.EnableEvents = True ' Znovu zapneme udalosti
End Sub

Požiadavky

    Microsoft Excel s podporou VBA (Office 365, Excel 2016 a vyššie).
    Povolené makrá: Pre správnu funkčnosť povoľte makrá pri otvorení súboru.

Licencia

Tento projekt je licencovaný pod licenciou MIT. Viac informácií nájdete v súbore LICENSE.

Tento README.md poskytuje detailné pokyny pre všetky aspekty projektu vrátane inštalácie, použitia a detailného popisu VBA kódu.
