# Macros
![excel](excel.png)

## Activer l'onglet développeur dans Excel
1. Ouvrir **Excel**
1. Clique droit sur le **ruban**
1. Chosir **Personnaliser**
1. Cocher la case **Développeur** dans la section de droite

## Avantages des Macros
- Intégré dans la suite *Office*
- Les macros permettent à des néophytes de faire de l'autotmatisation
- Le language *VBA*, bien que detesté par les programmeurs, est facile à comprendre pour un néophyte
- ~~C'est un beau language~~

## VBA
### Types VBA
|Type | Description|
|:-----:|------------|
|String | Contient du texte|
|Integer | Entier de -32 768 à 32 767|
|Single | Nombre réel|
|Boolean | Nombre réel|

### Variable en VBA
```VBA
Dim Toto As Integer
Toto = 5
```

### Cheat Sheet
La Cheat Sheet VBA est disponible [ici.](https://shiny.rstudio.com/images/rm-cheatsheet.png)

## Diagrame
Ce diagrame n'a pas de lien avec les reste du document.


```mermaid
graph TD

A(La machine à café ne fonctionne pas) --> B{La machine a du couant ?}
B -->|Non| C(Connecter la machine)
B -->|Oui| D{Manque d'eau ou de grains ?}
D -->|Oui| E{Avertissement filtre ?}
D -->|Non| F(Remplir eau et grains)
E -->|Oui| G(Remplacer ou nettoyer le filtre)
E -->|Non| H(Envoyer en réparation)

```
