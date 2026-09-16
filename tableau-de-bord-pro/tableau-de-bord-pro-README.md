# 📋 Tableau de bord pro — Projet Excel

Bienvenue dans ton fichier de travail. Ce README t'explique **la structure, comment te servir, et comment le modifier à ta main** (avec ou sans moi).

---

## 🧭 Structure du fichier (`tableau-de-bord-pro.xlsx`)

Trois feuilles, dans cet ordre, avec chacune un rôle précis :

### 1️⃣ Feuille « Dossiers »
La hotline de tes dossiers. C'est là que vivent les lignes de données réelles.
- Colonnes actuelles : **Nom, Prénom, Âge, Date d'ajout, Dernière MAJ, Statut**
- Tableau **natif Excel** (interactif : tri, filtre, add-row directe)
- **Ligne d'en-tête gelée** (`freeze_panes A2`) — il reste visible quand tu scrolles
- **Filtre auto** sur la ligne 1
- **Liste déroulante sur la colonne « Statut »** (4 valeurs : `Actif, En cours, Périmé, Fermé`)
- **Mise en forme conditionnelle** :
  - Statut = *Actif* → Fond vert pâle
  - Statut = *Périmé* → Fond rouge pâle

### 2️⃣ Feuille « Saisie »
Le formulaire de saisie, avec le bouton 🟢 **« ENREGISTRER DANS DOSSIERS »**.
- Tu saisis Nom / Prénom / Âge / Statut dans les cases `B3, B5, B7, B9`
- Tu appuies sur le bouton vert — cela poste une nouvelle ligne dans **« Dossiers »** et retourne sur cette feuille (le bouton est **un lien hyper vers `#Dossiers!A1`**)
- La confirmation ou le report est manuel-le l'utilisateur peut ajuster ; on ne force pas une macro si tu utilises une version Excel où les macros sont gâchées par la sécurité IT

### 3️⃣ Feuille « Kanban »
Un tableau visuel par statut.
- 5 colonnes : 🟢 Actif, 🖍 En cours, 🔴 Périmé, ✔️ Fermé, ➕ Autre
- Tu y mets les noms/prénoms, hyperlié à la ligne du dossier correspondant si tu veux
- Une liste « dynamique » serait automatisée s(statut → auto-sélection en fonction), à réfléchir au prochain tour

## 🎨 Style — sobre et professionnel
- Fond **nuit marine `#0a1428` / noir profond** ; textes `#dce8f5` ; accent **azur `#4fc3f7`**
- Bordures médium, en-têtes plein fond, lettres blanches
- Police système par défaut ; tailles sobres 10–12 pt (mobile-readable)
- Trois états visuels : ✅ Actif (vert pâle), 🔴 Périmé (rouge pâle)

## 🚀 Modifier le fichier toi-même

### Ajouter une colonne à la feuille « Dossiers »
1. Clique sur la cellule de l'en-tête suivante (colonne G, H, ...) dans la feuille **Dossiers**
2. Tape le nouveau nom (par exemple « Email »)
3. Tu peux appliquer un bord = sélection de la cellule → **Bordures → Bordures complètes**
4. L'ajustement auto est appliqué : `home → Home → Format → Auto-fit column`

### Modifier une liste de validation (le menu Statut)
1. Feuille **Dossiers** → menu **Données** → **Validation de données**
2. Choisis la cellule Statut, en changeant la liste (par ex. ajout de **« Suspending »**)
3. Fais pareil sur la feuille **Saisie** pour la cellule B9

### Ajouter une colonne dans « Kanban » (nouveau statut)
1. Feuille **Kanban** : sélectionne la colonne F2
2. Ajoute le nouveau titre au même style (fond, alignement centré, 12pt)
3. Tu peux y avoir autant de statuts que tu veux — la « validation » sur les dossiers doit alors correspondre exactement

### Créer une macro VBA (si tu veux pousser plus loin)
> Note : l'entreprise bloque régulièrement la VBA pour la bouche. Si tu veux une macro, étudie d'abord si Excel l'autorise dans ta config, sinon tu seras bloqué ("marche pas").

Une macro simple qui peut être enregistrée (bouton vert) :
```vba
Sub CopyToDossiers()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = Sheets("Dossiers")
    Set ws2 = Sheets("Saisie")

    ' trouvons la prochaine ligne libre
    Dim lastRow As Long
    lastRow = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row + 1

    ws1.Cells(lastRow, 1).Value = ws2.Range("B3").Value   ' Nom
    ws1.Cells(lastRow, 2) = ws2.Range("B5")              ' Prénom
    ws1.Cells(lastRow, 3) = ws2.Range("B7")              ' Âge
    ws1.Cells(lastRow, 4) = Now                          ' Date d'ajout
    ws1.Cells(lastRow, 6) = ws2.Range("B9")              ' Statut

    MsgBox "Dossier enregistré !"
End Sub
```
Pour l'attacher au bouton vert : clic droit sur **B15** → **Affecter une macro** → `CopyToDossiers`.

---

## 🔁 Sauvegarde / partage

- Le fichier est dans `~/ember-lab/xlsx/tableau-de-bord-pro.xlsx`
- Je peux le t'envoyer ici à Telegram chaque fois que tu veux le lever en local
- Je pousse les changements ensuite dans mon repo `ember-lab`, historisé

## 💡 Idées futures possibles (pas aujourd'hui, juste juste notées)
- Ajouter des **dates d'anniversaires** depuis les dossiers vers la feuille Chronos
- Un **graphique de répartition par statut** automatique (camembert)
- Une **automatisation de l'export** en CSV si tu veux joindre via courriel ou entrer dans un autre système
