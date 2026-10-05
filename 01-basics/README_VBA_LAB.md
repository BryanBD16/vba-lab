# 🎓 Premier Lab VBA - Guide Complet

## 📚 Concepts de Base du VBA

### Qu'est-ce que le VBA ?
**VBA (Visual Basic for Applications)** est un langage de programmation intégré dans Microsoft Office (surtout Excel). Il vous permet d'automatiser des tâches, créer des fonctions personnalisées, et manipuler des données.

---

## 🔑 Concepts Fondamentaux

### 1. **Les Variables**
Une variable est un conteneur pour stocker une valeur. Elle doit avoir un type spécifique.

```vba
Dim nom As String           ' Texte
Dim age As Integer          ' Nombre entier (sans décimale)
Dim prix As Double          ' Nombre décimal
Dim actif As Boolean        ' Vrai/Faux
Dim dateAuj As Date         ' Date
Dim plage As Range          ' Cellule ou plage Excel
```

**Exemple pratique :**
```vba
Sub ExempleVariables()
    Dim nomClient As String
    Dim montantVente As Double
    
    nomClient = "Jean Dupont"
    montantVente = 1250.50
    
    MsgBox "Client: " & nomClient & " | Montant: " & montantVente
End Sub
```

---

### 2. **Les Boucles**
Les boucles permettent de répéter une action plusieurs fois.

#### **Boucle FOR** (quand vous connaissez le nombre d'itérations)
```vba
For i = 1 To 10
    ' Code répété 10 fois
    MsgBox "Itération: " & i
Next i

' De 10 à 1 en décrémentant
For i = 10 To 1 Step -1
    MsgBox i
Next i
```

#### **Boucle WHILE** (tant qu'une condition est vraie)
```vba
Dim i As Integer
i = 1
While i <= 10
    MsgBox i
    i = i + 1
Wend
```

#### **Boucle FOR EACH** (pour parcourir une collection)
```vba
Dim feuille As Worksheet
For Each feuille In ThisWorkbook.Sheets
    MsgBox "Feuille: " & feuille.Name
Next feuille
```

---

### 3. **Les Conditions (IF)**
Permettent d'exécuter du code selon une condition.

```vba
Dim montant As Double
montant = 1500

If montant > 1000 Then
    MsgBox "Montant élevé!"
ElseIf montant > 500 Then
    MsgBox "Montant moyen"
Else
    MsgBox "Montant faible"
End If

' Plusieurs conditions avec AND/OR
If montant > 500 And montant < 2000 Then
    MsgBox "Dans la plage correcte"
End If

If montant < 100 Or montant > 5000 Then
    MsgBox "Montant en dehors des limites"
End If
```

---

### 4. **Subs vs Functions**

#### **Sub** - Exécute une action, ne retourne rien
```vba
Sub Saluer()
    MsgBox "Bonjour!"
End Sub

' Appel:
Call Saluer()
' ou simplement
Saluer()
```

#### **Function** - Retourne une valeur
```vba
Function Additionner(a As Integer, b As Integer) As Integer
    Additionner = a + b
End Function

Function Multiplier(x As Double, y As Double) As Double
    Multiplier = x * y
End Function

' Utilisation:
Sub TestFunctions()
    Dim resultat As Integer
    resultat = Additionner(5, 3)
    MsgBox "5 + 3 = " & resultat
End Sub
```

---

### 5. **Manipuler les Cellules et Plages**

```vba
' Accéder à une cellule
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Données")

ws.Range("A1").Value = "Bonjour"

' Accéder à une plage
ws.Range("A1:C10").Interior.Color = RGB(255, 255, 0)  ' Jaune

' Parcourir les cellules d'une plage
Dim cell As Range
For Each cell In ws.Range("A1:A100")
    MsgBox cell.Value
Next cell

' Trouver la dernière ligne avec données
Dim derniereLigne As Long
derniereLigne = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
MsgBox "Dernière ligne: " & derniereLigne
```

---

### 6. **Travailler avec les Données**

```vba
Sub LireEtAfficherDonnees()
    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long
    Dim nomClient As String
    Dim montant As Double
    
    Set ws = ThisWorkbook.Sheets("Données")
    derniereLigne = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Parcourir de la ligne 2 à la dernière (ligne 1 = en-têtes)
    For i = 2 To derniereLigne
        nomClient = ws.Cells(i, 2).Value      ' Colonne B
        montant = ws.Cells(i, 5).Value        ' Colonne E (Montant)
        
        If montant > 500 Then
            MsgBox nomClient & " a dépensé " & montant
        End If
    Next i
End Sub
```

---

### 7. **Gestion des Erreurs**

```vba
Sub AvecGestionErreur()
    On Error GoTo GestionErreur
    
    ' Votre code risqué ici
    Dim x As Integer
    x = 10 / 0  ' Erreur!
    
    Exit Sub
    
GestionErreur:
    MsgBox "Erreur détectée: " & Err.Description
    MsgBox "Numéro erreur: " & Err.Number
End Sub
```

---

## 📁 Structure du Lab

```
vba-lab/
├── README_VBA_LAB.md (ce fichier)
├── donnees_clients.xlsx (500 lignes de données)
├── exercices_vba.xlsm (macros à compléter)
└── SOLUTIONS/
    ├── Solution1_SommeColonne.bas
    ├── Solution2_FiltrerClients.bas
    ├── Solution3_RapportRegions.bas
    └── Solution4_ValidationDonnees.bas
```

---

## 🎯 TODO - Exercices Progressifs

### ✅ **EXERCICE 1 : Somme d'une Colonne** (⭐ Facile - 10 min)

**Objectif**: Créer une macro qui calcule la somme de tous les montants dans la colonne "Montant".

**À faire** :
- [ ] Ouvrir `exercices_vba.xlsm`
- [ ] Aller sur l'onglet "Sommaire"
- [ ] Créer une Sub `CalculerSomme()`
- [ ] Boucler sur les données (de ligne 2 à la dernière)
- [ ] Additionner la colonne E (Montant)
- [ ] Afficher le résultat avec `MsgBox`
- [ ] Cliquer le bouton "Exercice 1"

**Code à compléter** :
```vba
Sub CalculerSomme()
    Dim ws As Worksheet
    Dim somme As Double
    Dim i As Long
    Dim derniereLigne As Long
    
    Set ws = ThisWorkbook.Sheets("Données")
    derniereLigne = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    somme = 0
    
    ' TODO: Boucle FOR de 2 à derniereLigne
    '       Ajouter ws.Cells(i, 5).Value à somme
    
    ' TODO: MsgBox affichant "Somme totale: " & somme
    
End Sub
```

**Indices** :
- Colonne A = ID, B = Nom, C = Email, D = Région, E = Montant
- La boucle commence à i=2 (ligne 1 = en-têtes)
- Utilisez `ws.Cells(i, 5)` pour accéder à la colonne Montant

---

### ✅ **EXERCICE 2 : Filtrer les Données** (⭐⭐ Moyen - 20 min)

**Objectif**: Créer une macro qui crée une nouvelle feuille et y copie les clients avec montant > 500.

**À faire** :
- [ ] Créer une Sub `FiltrerClientsImportants()`
- [ ] Créer une nouvelle feuille "Clients Importants"
- [ ] Copier les en-têtes de la feuille "Données"
- [ ] Boucler et copier les lignes avec montant > 500
- [ ] Afficher le nombre de clients trouvés

**Concepts à pratiquer** :
- Boucles FOR
- Conditions IF
- Copie de données entre feuilles
- Création de feuilles

**Code à compléter** :
```vba
Sub FiltrerClientsImportants()
    Dim wsSource As Worksheet
    Dim wsNew As Worksheet
    Dim i As Long
    Dim ligneDestination As Long
    Dim montant As Double
    Dim derniereLigne As Long
    Dim nombreClients As Long
    
    Set wsSource = ThisWorkbook.Sheets("Données")
    
    ' TODO: Créer une nouvelle feuille nommée "Clients Importants"
    ' (Vérifier qu'elle n'existe pas déjà)
    
    ' TODO: Copier les en-têtes (ligne 1) dans la nouvelle feuille
    
    ligneDestination = 2
    nombreClients = 0
    
    ' TODO: Boucler sur les données
    ' TODO: Si montant > 500, copier la ligne
    ' TODO: Incrémenter nombreClients
    
    MsgBox nombreClients & " clients importants trouvés!"
    
End Sub
```

---

### ✅ **EXERCICE 3 : Créer un Rapport par Région** (⭐⭐ Moyen - 25 min)

**Objectif**: Créer un rapport qui compte le nombre de clients par région et le montant total.

**À faire** :
- [ ] Créer une Sub `GenererRapportRegions()`
- [ ] Créer une feuille "Rapport Régions"
- [ ] Compter les clients par région
- [ ] Calculer le montant total par région
- [ ] Afficher les résultats proprement

**Concept important : Collections/Dictionnaires**
```vba
' En VBA, on peut utiliser une Collection comme dictionnaire
Dim regions As Collection
Set regions = New Collection

' Ou utiliser un tableau associatif avec scripting
Set objDict = CreateObject("Scripting.Dictionary")
objDict.Add "Québec", 5
objDict.Add "Ontario", 3
```

**Données disponibles** :
Régions dans les données : Québec, Ontario, Manitoba, Colombie-Britannique, Alberta

---

### ✅ **EXERCICE 4 : Validation de Données** (⭐⭐⭐ Avancé - 30 min)

**Objectif**: Créer des fonctions de validation et marquer les données invalides.

**À faire** :
- [ ] Créer une Function `ValiderEmail(email As String) As Boolean`
- [ ] Créer une Function `ValiderMontant(montant As Double) As Boolean`
- [ ] Créer une Sub `ValiderToutesLes Donnees()`
- [ ] Parcourir toutes les données
- [ ] Marquer les lignes invalides en rouge
- [ ] Afficher un rapport

**Règles de validation** :
- Email : doit contenir "@" et un point après "@"
- Montant : doit être entre 0 et 10000

**Code à compléter** :
```vba
Function ValiderEmail(email As String) As Boolean
    ' TODO: Vérifier que email contient "@"
    ' TODO: Vérifier que après "@" il y a un "."
    ' Retourner True si valide, False sinon
End Function

Function ValiderMontant(montant As Double) As Boolean
    ' TODO: Vérifier que montant >= 0 AND montant <= 10000
    ' Retourner True si valide, False sinon
End Function

Sub ValiderToutesLesDonnees()
    Dim ws As Worksheet
    Dim i As Long
    Dim derniereLigne As Long
    Dim email As String
    Dim montant As Double
    Dim erreurs As Long
    
    Set ws = ThisWorkbook.Sheets("Données")
    derniereLigne = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    erreurs = 0
    
    ' TODO: Parcourir toutes les données
    ' TODO: Vérifier email et montant
    ' TODO: Si invalide, marquer la ligne en rouge
    ' TODO: Compter les erreurs
    
    MsgBox erreurs & " erreurs détectées!"
    
End Sub
```

---

## 🛠️ Comment Commencer

### 1. **Ouvrir Excel**
Assurez-vous d'avoir Excel avec accès au VBA.

### 2. **Accéder à l'Éditeur VBA**
- Appuyer sur **Alt + F11** 
- Ou : Onglet "Développeur" → "Visual Basic"
- Ou : Outils → Macros → Éditer les macros

### 3. **Insérer un Module**
```
Clic droit sur le document → Insérer → Module
```

### 4. **Copier le Code VBA**
Copier-coller le code des solutions dans le module.

### 5. **Exécuter une Macro**
- **Alt + F8** pour ouvrir la liste des macros
- Sélectionner la macro et cliquer "Exécuter"
- Ou attacher un bouton au formulaire

### 6. **Déboguer**
- Utiliser **MsgBox** pour afficher les valeurs
- Ajouter des points d'arrêt (clic sur la marge grise)
- Appuyer sur **F8** pour exécuter ligne par ligne

---

## 📊 Données Disponibles

Le fichier `donnees_clients.xlsx` contient :
- **500 lignes** de données clients
- **Colonnes** :
  - A : ID (1-500)
  - B : Nom (Noms aléatoires)
  - C : Email (format: prenom.nom@domain.com)
  - D : Région (Québec, Ontario, Manitoba, Colombie-Britannique, Alberta)
  - E : Montant (100-10000)
  - F : Date (dates aléatoires 2023-2026)

**Parfait pour pratiquer** : boucles, conditions, copie de données !

---

## 💡 Conseils et Bonnes Pratiques

### ✅ **Nommage des Variables**
```vba
' ❌ Mauvais
Dim x As String
Dim a As Integer

' ✅ Bon
Dim nomClient As String
Dim montantTotal As Integer
Dim wsSource As Worksheet
```

### ✅ **Commentaires Utiles**
```vba
' Ceci est un commentaire
' Expliquez la logique de votre code

' Boucler sur tous les clients
For i = 2 To derniereLigne
    ' Récupérer le montant
    montant = ws.Cells(i, 5).Value
Next i
```

### ✅ **Performance pour Grandes Données**
```vba
' Avant de traiter beaucoup de données
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

' Votre code ici
For i = 2 To 10000
    ' ...
Next i

' Réactiver après
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
```

### ✅ **Gestion des Erreurs**
```vba
Sub AvecSecurite()
    On Error GoTo Erreur
    
    ' Votre code
    Dim x As Integer
    x = 10 / 0  ' Erreur!
    
    Exit Sub
    
Erreur:
    MsgBox "Une erreur s'est produite: " & Err.Description
End Sub
```

---

## 📋 Plan d'Apprentissage Recommandé

| Jour | Exercices | Durée | Concepts |
|------|-----------|-------|----------|
| 1 | Exercice 1 | 15 min | Variables, boucles FOR |
| 2 | Exercice 2 | 30 min | Conditions IF, manipulation de feuilles |
| 3 | Exercice 3 | 30 min | Collections, boucles complexes |
| 4 | Exercice 4 | 40 min | Functions, validation |
| 5 | Mini-projet personnel | 1h+ | Réunir tous les concepts |

---

## 🎨 Mini-Projet Personnel (après les exercices)

**Idée 1 : Facture Automatique**
- Lire une liste de produits
- Calculer le total
- Générer une facture formatée

**Idée 2 : Suivi des Ventes**
- Importer les données clients
- Créer un résumé par région
- Générer un rapport mensuel

**Idée 3 : Nettoyage de Données**
- Supprimer les doublons
- Formater les emails
- Valider les montants

---

## 🔗 Ressources Utiles

- **Microsoft VBA Documentation** : https://docs.microsoft.com/office/vba/
- **Excel VBA Reference** : https://docs.microsoft.com/office/vba/api/overview/excel
- **Stack Overflow** : Chercher "vba" ou "excel vba"

---

**Bonne chance dans votre apprentissage du VBA ! 🚀**

Des questions ? Des exercices trop faciles/difficiles ? N'hésitez pas à me contacter !
