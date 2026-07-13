# ============================================================
#  CONVERTISSEUR DE COURS → WORD  (LPA / DPFC Côte d'Ivoire)
#  Google Colab — python-docx  — VERSION 4 (style modèle v2)
#
#  UTILISATION :
#  1. Installe : !pip install python-docx -q
#  2. Colle ton cours dans la variable COURS ci-dessous
#  3. Exécute → téléchargement automatique
#
#  FORMAT DU COURS (inchangé) :
#  ════...════  en-tête  ════...════
#  TITRE_LECON: ...
#  COMPETENCE: ...
#  PROFESSEUR: ...
#  ORGANISME: ...
#  SOMMAIRE
#  ────...────
#  SECTION N — ...
#  [CAPSULE DE REPRISE — ...]
#  [ANCRAGE IVOIRIEN]
#  ▶ HARKNESS N — ...
#  Phase N · ...
#  ~~~  texte  ---TRADUCTION---  ~~~
#  ┌─ À RETENIR — ... ─┐ ... └─┘
#  ⚠ ATTENTION — ...
#  ◆ DÉFINITION DPFC — ...
#  ✎ TRACE ÉCRITE
#  ~ ANALOGIE CI — N
#  SYNTHESE_DEBUT ... SYNTHESE_FIN
#  ◉ EXERCICE GUIDÉ N — ...
#  ◎ EXERCICE N — ...
#  ◈ DEVOIR MAISON — ...
#  ✔ CORRIGÉ — ...
#  [SCHÉMA N — ...]
# ============================================================

COURS = """════════════════════════════════════════════════════════
  ANGLAIS LV1 · Leçon 14 — Village life (Vie au village)
  Niveau : 3ème | Compétence : C5 — Écoute simple
  Professeur : Evelyne ASSI
  Programme : DPFC/MENET-FP CI 2025-2026
════════════════════════════════════════════════════════

TITRE_LECON: Village life — Vie au village
COMPETENCE: Compétence C5 · Écoute simple · 3ème
PROFESSEUR: Professeur Evelyne ASSI  ·  Coefficient 2
ORGANISME: DPFC / MENET-FP Côte d'Ivoire  ·  2025-2026

SOMMAIRE
  1. Avant la leçon
  2. La leçon
  3. Après la leçon
  4. Corrigés complets

────────────────────────────────────────────────────────
SECTION 1 — AVANT LA LEÇON
────────────────────────────────────────────────────────

[CAPSULE DE REPRISE — Leçon 13 : City life]

  5 points clés de la leçon précédente :
  Point 1 · La vie en ville (city life) a des avantages : emplois (jobs), services, modernité, divertissements.
  Point 2 · La vie en ville a des inconvénients : bruit (noise), pollution, embouteillages (traffic jams), loyers chers (expensive rent).
  Point 3 · "There is" s'utilise avec un nom SINGULIER ; "There are" avec un nom PLURIEL.
  Point 4 · Comparatif des adjectifs courts (1 syllabe) : adj + -ER + than → big → bigger than.
  Point 5 · Comparatif des adjectifs longs (2+ syllabes) : more + adj + than → expensive → more expensive than. ⚠ Jamais "more bigger" — erreur grave.

  3 questions flash :
  Q1 · Complète : "There ________ many schools in Abidjan." → There ARE many schools in Abidjan.
  Q2 · Forme le comparatif de "hot". → hotter than (CVC → double T)
  Q3 · Traduis : "The city is more polluted than the village." → La ville est plus polluée que le village.

  Lien avec la leçon du jour : En Leçon 13, tu as décrit la vie en VILLE. Aujourd'hui tu décris son opposé : la vie au VILLAGE.

────────────────────────────────────────────────────────

[ANCRAGE IVOIRIEN]

  En Côte d'Ivoire, plus de 40 % de la population vit encore en zone rurale. Des villages comme Tiébissou, Daoukro, Katiola ou Boundiali concentrent des communautés agricoles qui cultivent le cacao, le café, l'igname et le riz.
  Lien DPFC : Environnement · agriculture · civisme · valorisation du patrimoine culturel ivoirien.

────────────────────────────────────────────────────────

▶ HARKNESS 1 — Qu'est-ce qu'un village ?
  Q : "How is a village different from a city?"
  Guide : Piste 1 · Pense au nombre d'habitants. | Piste 2 · Pense aux activités principales. | Piste 3 · Pense aux services disponibles.
  Corrigé : A village is a small community where fewer people live. People work in agriculture. Unlike cities, villages have fewer services.

▶ HARKNESS 2 — Les avantages de la vie au village
  Q : "Why do some people prefer village life to city life?"
  Guide : Piste 1 · Tranquillité et nature. | Piste 2 · Liens communautaires. | Piste 3 · Coût de la vie.
  Corrigé : Village life is quieter and less stressful. People live close to nature and eat fresh food. Community life is stronger.

▶ HARKNESS 3 — Le superlatif
  Q : "How do you say that something is THE MOST or THE LEAST in English?"
  Guide : Piste 1 · Différence bigger / the biggest. | Piste 2 · Superlatif adjectif long. | Piste 3 · Article toujours utilisé.
  Corrigé : Short adjectives : the + adj + -EST → the biggest. Long adjectives : the most + adj → the most beautiful. ALWAYS use "the".

────────────────────────────────────────────────────────
SECTION 2 — LA LEÇON
────────────────────────────────────────────────────────

Phase 1 · Présentation / Prérequis

  Titre : Village life (La vie au village)
  Compétence : C5 — Écoute simple
  Professeur Evelyne ASSI : "La vie au village, c'est la vie de millions d'Ivoiriens. Parlons-en en anglais !"

  Objectifs de la leçon
  Décrire la vie au village en anglais
  Connaître le vocabulaire des activités et lieux ruraux
  Utiliser le superlatif correctement

  Prérequis nécessaires
  Vocabulaire de la ville (Leçon 13)
  Le comparatif (Leçon 13)
  Le Present Simple (habitudes)

────────────────────────────────────────────────────────

Phase 2 · Découverte guidée

~~~
My name is Koné. I come from Katiola, a small village in the north of Côte d'Ivoire.
The village is the quietest place I know. The air is the freshest and the nights are the most peaceful.
However, the nearest hospital is fifty kilometres away, and there is no Internet connection.
---TRADUCTION---
Je m'appelle Koné. Je viens de Katiola, un petit village du nord de la Côte d'Ivoire.
Le village est l'endroit le plus calme que je connaisse. L'air est le plus frais et les nuits sont les plus paisibles.
Cependant, l'hôpital le plus proche est à cinquante kilomètres, et il n'y a pas de connexion Internet.
~~~

┌─ À RETENIR — LE SUPERLATIF ─────────────────────────┐
  Le superlatif exprime le DEGRÉ MAXIMUM ou MINIMUM parmi trois éléments ou plus.
  On utilise TOUJOURS "the" devant le superlatif.
  Adjectif COURT (1 syllabe) → the + adj + -EST : quiet → the quietest · fresh → the freshest
  ⚠ CVC : big → the biggest (double G) · hot → the hottest (double T)
  Adjectif LONG (2 syllabes+) → the most + adj : peaceful → the most peaceful
  Forme négative : the least + adj → the least expensive
└──────────────────────────────────────────────────────┘

⚠ ATTENTION — Pièges et faux amis
  · "village" se prononce /ˈvɪl.ɪdʒ/ — VI-lidj PAS "vilaghe".
  · "fetch" = aller chercher ET ramener — PAS juste "chercher".
  · "crop" = culture agricole / récolte.

────────────────────────────────────────────────────────

Phase 3 · Schémas textuels

[SCHÉMA 1 — Carte mentale : Village Life Vocabulary]
  Centre : VILLAGE LIFE / VIE AU VILLAGE
  PLACES : field · river · mango tree · well · path
  ACTIVITIES : grow crops · fetch water · cook · walk to school
  PROBLEMS : no hospital · no Internet · few jobs · limited electricity
  ADVANTAGES : quiet · fresh air · peaceful · cheap · strong community
  Note générateur : Canva (mind map) ou MindMeister

[SCHÉMA 2 — Tableau : Comparatif vs Superlatif]
  quiet | quieter than | the quietest
  big | bigger than | the biggest (CVC → double G)
  peaceful | more peaceful than | the most peaceful
  expensive | more expensive than | the most expensive
  ⚠ TOUJOURS "the" devant le superlatif
  Note générateur : Canva (tableau) ou draw.io

────────────────────────────────────────────────────────

Phase 4 · Définitions DPFC

◆ DÉFINITION DPFC — Village
  A village is a small community in a rural area where people live, usually working in agriculture and maintaining strong traditional and community ties.
  (Un village est une petite communauté en zone rurale où les gens vivent, travaillant généralement dans l'agriculture.)

◆ DÉFINITION DPFC — Superlative (superlatif)
  The superlative is a grammatical form used to express the highest or lowest degree of a quality among three or more elements. Always used with "the".
  (Le superlatif exprime le degré le plus élevé ou le plus bas parmi trois éléments ou plus. Toujours avec "the".)

◆ DÉFINITION DPFC — Crop (culture agricole)
  A crop is a plant grown by farmers in fields for food, such as yam, maize, cocoa or rice.
  (Une culture agricole est une plante cultivée dans des champs pour se nourrir.)

────────────────────────────────────────────────────────

Phase 5 · Trace écrite + Analogies CI

✎ TRACE ÉCRITE (à recopier dans le cahier)

  LEÇON 14 — VILLAGE LIFE (La vie au village)
  Compétence C5 · Écoute simple · Anglais LV1 · 3ème

  I. VOCABULAIRE DE LA VIE AU VILLAGE
  Lieux : field = champ · river = rivière · well = puits · mango tree = manguier
  Activités : grow crops = cultiver · fetch water = aller chercher de l'eau · cook = cuisiner
  Avantages : quiet = calme · peaceful = paisible · fresh air = air frais · cheap = bon marché
  Inconvénients : no hospital · no Internet · few jobs · limited electricity

  II. LE SUPERLATIF
  Adjectif COURT → the + adj + -EST : quiet → the quietest · fresh → the freshest
  ⚠ CVC : big → the biggest (double G) · hot → the hottest (double T)
  Adjectif LONG → the most + adj : peaceful → the most peaceful
  ⚠ RÈGLE ABSOLUE : TOUJOURS "the" devant le superlatif.

  III. COMPARATIF vs SUPERLATIF
  Comparatif (2 éléments) : "The village is quieter THAN the city."
  Superlatif (degré absolu) : "The village is THE quietest place."

~ ANALOGIE CI — 1
  Le superlatif, c'est comme proclamer le champion d'une compétition de lutte à Bouaké : il n'y en a qu'UN qui est "the strongest". Le comparatif, c'est juste dire qui est plus fort dans un match 1-contre-1.

~ ANALOGIE CI — 2
  "Fetch water" au village, c'est aller au marigot avec une bassine et rentrer avec cette bassine pleine. "Fetch" = aller ET revenir avec quelque chose — un aller-retour complet.

────────────────────────────────────────────────────────

Phase 5b · Synthèse

SYNTHESE_DEBUT
  5 points essentiels à retenir
  1. La vie au village est caractérisée par l'agriculture, la nature, la communauté et la tranquillité.
  2. Le superlatif exprime le degré maximum parmi trois éléments ou plus.
  3. Superlatif adjectif court : the + adj + -EST.
  4. Superlatif adjectif long : the most + adj.
  5. "The" est TOUJOURS présent avant le superlatif.

  5 mots clés
  · Field = champ (terrain cultivé par les agriculteurs)
  · Crop = récolte / culture (plante cultivée pour se nourrir)
  · Peaceful = paisible (calme et tranquille)
  · Community = communauté (groupe partageant le même lieu)
  · Fetch = aller chercher et ramener

  3 erreurs fréquentes + correction
  ✗ "Village is quietest place." → ✓ The village is THE quietest place.
  ✗ "The most quietest" → ✓ the quietest OU the most peaceful — jamais les deux.
  ✗ Confondre comparatif et superlatif → ✓ Comparatif = 2 éléments / Superlatif = degré absolu.
SYNTHESE_FIN

────────────────────────────────────────────────────────

Phase 6 · Exercices guidés

◉ EXERCICE GUIDÉ 1 — Le superlatif
  Notion ciblée : formation du superlatif
  Énoncé : Transforme ces phrases comparatives en superlatives.
  1. The mango tree is taller than the other trees. → The mango tree is ________ tree in the village.
  2. Village air is fresher than city air. → Village air is ________ air in the country.
  3. The river is more beautiful than the well. → The river is ________ place in the village.
  MÉTHODE
  Étape 1 · Identifie l'adjectif dans la phrase comparative.
  Étape 2 · Court ou long ? tall → 1 syl. → the tallest · beautiful → 3 syl. → the most beautiful
  Étape 3 · Ajoute "the" obligatoirement.
  ✔ Résultats :
  1. The mango tree is the tallest tree in the village.
  2. Village air is the freshest air in the country.
  3. The river is the most beautiful place in the village.

◉ EXERCICE GUIDÉ 2 — Décrire la vie au village
  Notion ciblée : Present Simple + vocabulaire rural + superlatif en contexte
  Énoncé : Rédige un paragraphe de 4 phrases sur la famille Diabaté de Boundiali.
  MÉTHODE
  Étape 1 · Présente la famille : "The Diabaté family lives in Boundiali. They grow yam and maize."
  Étape 2 · Routine matinale : "Every morning, the children walk to school."
  Étape 3 · Soirée : "In the evening, the family sits together under the mango tree."
  Étape 4 · Superlatif : "For Koné, the village is the most peaceful place in the world."
  ✔ Paragraphe complet : "The Diabaté family lives in Boundiali. They grow yam and maize. Every morning, the children walk to school. In the evening, the family sits together under the mango tree. For Koné, the village is the most peaceful place in the world."

────────────────────────────────────────────────────────
SECTION 3 — APRÈS LA LEÇON
────────────────────────────────────────────────────────

◎ EXERCICE 1 — Vocabulaire de la vie au village
  Notions travaillées : village life vocabulary
  Associe chaque mot anglais à sa traduction française.
  Colonne A : 1. field  2. fetch water  3. crop  4. mango tree  5. peaceful
  Colonne B : a. paisible  b. manguier  c. champ  d. aller chercher de l'eau  e. récolte
  GUIDE
  Étape 1 · Relis la trace écrite. | Étape 2 · Cherche chaque traduction. | Étape 3 · Note : 1→c, etc.

◎ EXERCICE 2 — Present Simple : habitudes rurales
  Notions travaillées : Present Simple + vocabulaire rural
  Mets les verbes au Present Simple. Traduis chaque phrase.
  1. Farmers (grow) ________ yam and maize every season.
  2. The woman (fetch) ________ water from the river.
  3. Children (walk) ________ to school every morning.
  GUIDE
  Étape 1 · Identifie le sujet. | Étape 2 · Singulier he/she/it → verbe + -S. | Étape 3 · Écris et traduis.

◎ EXERCICE 3 — Formation du superlatif
  Notions travaillées : règles du superlatif
  Donne le superlatif. Précise la règle.
  1. small → ________ | 2. beautiful → ________ | 3. hot → ________ | 4. safe → ________ | 5. difficult → ________
  GUIDE
  Étape 1 · Compte les syllabes. | Étape 2 · Vérifie CVC si nécessaire. | Étape 3 · Applique la règle, n'oublie pas "the".

◎ EXERCICE 4 — Comparer ville et village
  Notions travaillées : comparatif + superlatif en contexte
  Écris 4 phrases comparant ou décrivant au superlatif la ville et le village.
  Adjectifs : noisy · cheap · modern · exciting
  GUIDE
  Étape 1 · Court ou long ? | Étape 2 · Comparatif. | Étape 3 · Superlatif. | Étape 4 · Deux phrases par adjectif.

◎ EXERCICE 5 — Rédiger un texte sur la vie au village
  Notions travaillées : cohérence textuelle + Present Simple + superlatif
  Aminata vit à Daoukro. Écris un paragraphe de 70 à 90 mots décrivant sa vie.
  Inclure : 3 phrases au Present Simple · 2 superlatifs · 4 mots de vocabulaire rural · 1 avantage ET 1 inconvénient.
  GUIDE
  Étape 1 · Présente Aminata. | Étape 2 · Habitudes (Present Simple). | Étape 3 · Avantage au superlatif. | Étape 4 · Inconvénient. | Étape 5 · Conclusion.

◈ DEVOIR MAISON — Village Life vs City Life
  Durée conseillée : 35 minutes
  CONTEXTE :
  Ton professeur t'a demandé d'écrire un texte comparatif pour le journal du lycée.
  ÉNONCÉ :
  Rédige un texte de 100 à 120 mots comparant la vie au village et la vie en ville en Côte d'Ivoire.
  STRUCTURE :
  1. Introduction : présente les deux milieux.
  2. Paragraphe village : avantages + inconvénients.
  3. Paragraphe ville : avantages + inconvénients.
  4. Conclusion : quel milieu préfères-tu et pourquoi ?
  CONTRAINTES OBLIGATOIRES :
  · Minimum 3 phrases au Present Simple
  · Minimum 2 comparatifs (avec "than")
  · Minimum 2 superlatifs (avec "the")
  · Minimum 5 mots de vocabulaire variés
  · Au moins 2 noms de lieux ivoiriens réels
  GUIDE PAS À PAS (sans corrigé)
  Étape 1 · Introduction (2 phrases).
  Étape 2 · Paragraphe village (3-4 phrases).
  Étape 3 · Paragraphe ville (3-4 phrases).
  Étape 4 · Conclusion (2 phrases).
  Étape 5 · Compte tes mots (100 à 120).
  Étape 6 · Relis et vérifie.

────────────────────────────────────────────────────────
SECTION 4 — CORRIGÉS COMPLETS
────────────────────────────────────────────────────────

✔ CORRIGÉ — DEVOIR MAISON
  TITRE PROPOSÉ : "Village Life and City Life in Côte d'Ivoire"
  In Côte d'Ivoire, some people live in big cities like Abidjan, while others live in small villages like Daoukro or Katiola.
  Village life is the most peaceful way of living. Farmers grow yam and maize in their fields every day. The air is fresher and the community is stronger than in the city. However, there are no hospitals nearby and there is no Internet.
  City life is more exciting than village life. There are many jobs and modern hospitals in Abidjan. But the city is the most stressful place to live.
  In my opinion, I prefer village life because it is the most peaceful and the cheapest way to live. (117 mots)
  TRADUCTION : En Côte d'Ivoire, certaines personnes vivent dans de grandes villes comme Abidjan, tandis que d'autres vivent dans de petits villages. La vie au village est le mode de vie le plus paisible. Cependant, il n'y a pas d'hôpitaux. La vie en ville est plus excitante mais plus stressante.

✔ CORRIGÉ — EXERCICE 1
  1→c (field = champ) · 2→d (fetch water = aller chercher de l'eau) · 3→e (crop = récolte) · 4→b (mango tree = manguier) · 5→a (peaceful = paisible)

✔ CORRIGÉ — EXERCICE 2
  1. Farmers grow yam and maize every season. (they → pas de -S)
  2. The woman fetches water from the river. (she → fetch + -ES)
  3. Children walk to school every morning. (they → pas de -S)

✔ CORRIGÉ — EXERCICE 3
  1. small → the smallest (court 1 syl. → the + -EST)
  2. beautiful → the most beautiful (long 3 syl. → the most + adj)
  3. hot → the hottest (CVC [h-o-t] → double T)
  4. safe → the safest (court en -E → ajoute -ST)
  5. difficult → the most difficult (long 3 syl. → the most + adj)

✔ CORRIGÉ — EXERCICE 4
  noisy → The city is noisier than the village. / Abidjan is the noisiest city I know.
  cheap → The village is cheaper than the city. / The village is the cheapest place to live.
  modern → The city is more modern than the village. / Abidjan is the most modern city in CI.
  exciting → City life is more exciting than village life. / The city is the most exciting place for young people.

✔ CORRIGÉ — EXERCICE 5
  "My name is Aminata and I live in Daoukro. Every day, my mother fetches water from the river and cooks traditional meals. My father grows yam and maize in the fields. Daoukro is the most peaceful place I know — the air is the freshest. However, there is no hospital nearby and there is no Internet. For me, village life is the best."
  ✔ Present Simple : fetches, cooks, grows ✔ · ✔ Superlatifs : the most peaceful, the freshest ✔ · ✔ Vocabulaire : river, fields, yam, maize ✔

════════════════════════════════════════════════════════
  Prof. Evelyne ASSI · Anglais LV1 · 3ème
  DPFC/MENET-FP CI · 2025-2026
  "Speak even if it's not perfect — silence teaches nothing."
  (Parle même si c'est imparfait — le silence n'apprend rien.)
════════════════════════════════════════════════════════
"""

# ============================================================
#  GÉNÉRATEUR — VERSION 4 (style modèle de référence v2)
# ============================================================

import re
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── COULEURS ─────────────────────────────────────────────────
C_BLEU_TITRE  = RGBColor(0x18, 0x40, 0x82)
C_OR_TITRE    = RGBColor(0xC9, 0x92, 0x25)
C_GRIS_TEXTE  = RGBColor(0x23, 0x23, 0x23)
C_GRIS_CLAIR  = RGBColor(0x55, 0x55, 0x55)
C_VERT_OK     = RGBColor(0x2A, 0x7E, 0x5B)
C_ROUGE_ERR   = RGBColor(0xA6, 0x2D, 0x2D)
C_ORANGE_ANAL = RGBColor(0xC0, 0x54, 0x1A)
C_BLANC       = RGBColor(0xFF, 0xFF, 0xFF)

# ── FONDS ────────────────────────────────────────────────────
BG_SECTION    = RGBColor(0x18, 0x3E, 0x82)
BG_CAPSULE    = RGBColor(0xF0, 0xEB, 0xF8)
BG_ANCRAGE    = RGBColor(0xE2, 0xF0, 0xD9)
BG_HARKNESS   = RGBColor(0xEA, 0xF2, 0xF8)
BG_DECOUVERTE = RGBColor(0xEA, 0xF2, 0xF8)
BG_RETENIR    = RGBColor(0xE2, 0xF0, 0xD9)
BG_ATTENTION  = RGBColor(0xFF, 0xF2, 0xCC)
BG_DEF        = RGBColor(0xEA, 0xF2, 0xF8)
BG_TRACE      = RGBColor(0xED, 0xF4, 0xFB)
BG_ANALOGIE   = RGBColor(0xFE, 0xF0, 0xE7)
BG_SYNTHESE   = RGBColor(0xD9, 0xE2, 0xF3)
BG_EXO_GUIDE  = RGBColor(0xF5, 0xF7, 0xFA)
BG_EXO_ENT    = RGBColor(0xFA, 0xFA, 0xFA)
BG_DEVOIR     = RGBColor(0xFF, 0xF9, 0xEE)
BG_CORRIGE    = RGBColor(0xF8, 0xFB, 0xF7)
BG_SCHEMA     = RGBColor(0xFA, 0xFA, 0xFA)
BG_SOMMAIRE   = RGBColor(0xEE, 0xF3, 0xFA)


def _hex(c):
    return f"{c[0]:02X}{c[1]:02X}{c[2]:02X}"


def set_cell_bg(cell, color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for e in tcPr.findall(qn('w:shd')):
        tcPr.remove(e)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), _hex(color))
    tcPr.append(shd)


def remove_cell_borders(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "none")
        b.set(qn("w:sz"), "0")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "auto")
        tcBorders.append(b)
    tcPr.append(tcBorders)


def add_cell_border_bottom(cell, color="184082", sz="4"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), sz)
    bottom.set(qn("w:space"), "0")
    bottom.set(qn("w:color"), color)
    tcBorders.append(bottom)
    tcPr.append(tcBorders)


def add_cell_border_top(cell, color="184082", sz="4"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), sz)
    top.set(qn("w:space"), "0")
    top.set(qn("w:color"), color)
    tcBorders.append(top)
    tcPr.append(tcBorders)


doc = Document()
for sec in doc.sections:
    sec.top_margin    = Cm(1.8)
    sec.bottom_margin = Cm(1.8)
    sec.left_margin   = Cm(2.0)
    sec.right_margin  = Cm(2.0)


# ── EN-TÊTE ──────────────────────────────────────────────────
def add_header(lecon_info):
    section = doc.sections[0]
    header = section.header
    header.is_linked_to_previous = False
    tbl = header.add_table(rows=1, cols=3, width=Cm(17))
    tbl.style = "Table Grid"
    cells = tbl.rows[0].cells
    for cell in cells:
        remove_cell_borders(cell)
        add_cell_border_bottom(cell)
    p0 = cells[0].paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r0 = p0.add_run("LPA · 3ème")
    r0.font.size = Pt(8); r0.font.color.rgb = C_GRIS_CLAIR
    p1 = cells[1].paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p1.add_run(lecon_info)
    r1.bold = True; r1.font.size = Pt(8); r1.font.color.rgb = C_BLEU_TITRE
    p2 = cells[2].paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r2 = p2.add_run("Anglais LV1")
    r2.font.size = Pt(8); r2.font.color.rgb = C_GRIS_CLAIR


# ── PIED DE PAGE ─────────────────────────────────────────────
def add_footer(prof, organisme):
    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False
    tbl = footer.add_table(rows=1, cols=3, width=Cm(17))
    tbl.style = "Table Grid"
    cells = tbl.rows[0].cells
    for cell in cells:
        remove_cell_borders(cell)
        add_cell_border_top(cell)
    p0 = cells[0].paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r0 = p0.add_run(prof); r0.font.size = Pt(8); r0.font.color.rgb = C_GRIS_CLAIR
    p1 = cells[1].paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p1.add_run(organisme); r1.font.size = Pt(8); r1.font.color.rgb = C_GRIS_CLAIR
    p2 = cells[2].paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r2 = p2.add_run("Page "); r2.font.size = Pt(8); r2.font.color.rgb = C_GRIS_CLAIR
    fldChar1 = OxmlElement("w:fldChar"); fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText"); instrText.text = "PAGE"
    fldChar2 = OxmlElement("w:fldChar"); fldChar2.set(qn("w:fldCharType"), "end")
    r_p = p2.add_run(); r_p.font.size = Pt(8); r_p.font.color.rgb = C_GRIS_CLAIR
    r_p._r.append(fldChar1); r_p._r.append(instrText); r_p._r.append(fldChar2)


# ── PAGE DE GARDE ────────────────────────────────────────────
def add_page_garde(titre, competence, professeur, organisme, lecon_num):
    # Grand LPA
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(50)
    p.paragraph_format.space_after  = Pt(4)
    r = p.add_run("LPA")
    r.bold = True; r.font.size = Pt(60); r.font.color.rgb = C_BLEU_TITRE
    # Sous-titre
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after  = Pt(16)
    r2 = p2.add_run("Lycée Personnel Autonome")
    r2.italic = True; r2.font.size = Pt(13); r2.font.color.rgb = C_GRIS_CLAIR
    # Ligne bleue
    p_sep = doc.add_paragraph()
    p_sep.paragraph_format.space_before = Pt(0)
    p_sep.paragraph_format.space_after  = Pt(14)
    pPr = p_sep._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), '16')
    bottom.set(qn('w:space'), '1'); bottom.set(qn('w:color'), '184082')
    pBdr.append(bottom); pPr.append(pBdr)
    # Matière · Leçon · Niveau
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_before = Pt(0)
    p3.paragraph_format.space_after  = Pt(6)
    r3 = p3.add_run(f"Anglais LV1  ·  {lecon_num}  ·  Classe de 3ème")
    r3.bold = True; r3.font.size = Pt(14); r3.font.color.rgb = C_BLEU_TITRE
    # Grand titre
    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4.paragraph_format.space_before = Pt(0)
    p4.paragraph_format.space_after  = Pt(6)
    r4 = p4.add_run(titre)
    r4.bold = True; r4.font.size = Pt(32); r4.font.color.rgb = C_OR_TITRE
    # Compétence italique
    p5 = doc.add_paragraph()
    p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p5.paragraph_format.space_before = Pt(0)
    p5.paragraph_format.space_after  = Pt(30)
    r5 = p5.add_run(competence)
    r5.italic = True; r5.font.size = Pt(12); r5.font.color.rgb = C_GRIS_CLAIR
    # Professeur
    p6 = doc.add_paragraph()
    p6.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p6.paragraph_format.space_before = Pt(0)
    p6.paragraph_format.space_after  = Pt(6)
    r6 = p6.add_run(professeur)
    r6.font.size = Pt(12); r6.font.color.rgb = C_GRIS_TEXTE
    # Organisme
    p7 = doc.add_paragraph()
    p7.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p7.paragraph_format.space_before = Pt(0)
    p7.paragraph_format.space_after  = Pt(6)
    r7 = p7.add_run(organisme)
    r7.font.size = Pt(11); r7.font.color.rgb = C_GRIS_CLAIR
    # Saut de page
    doc.add_page_break()


# ── SOMMAIRE ─────────────────────────────────────────────────
def add_sommaire(sommaire_lines, sections_phases):
    """
    sommaire_lines : lignes du cours (1. Avant la leçon, etc.)
    sections_phases : dict pour enrichir avec les phases
    """
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, BG_SOMMAIRE)
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in [("top","80"),("bottom","80"),("left","140"),("right","140")]:
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"), val); m.set(qn("w:type"), "dxa")
        tcMar.append(m)
    tcPr.append(tcMar)
    # Titre Sommaire
    p0 = cell.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(8)
    r0 = p0.add_run("Sommaire")
    r0.bold = True; r0.font.size = Pt(16); r0.font.color.rgb = C_BLEU_TITRE

    # Mapping sections → sous-points enrichis
    enriched = {
        "1": ["· Capsule de reprise", "· Ancrage ivoirien", "· Questions Harkness"],
        "2": ["· Phase 1 — Présentation & prérequis", "· Phase 2 — Découverte guidée",
              "· Phase 3 — Schémas textuels", "· Phase 4 — Définitions DPFC",
              "· Phase 5 — Trace écrite + Analogies CI", "· Phase 5b — Synthèse",
              "· Phase 6 — Exercices guidés"],
        "3": ["· Exercices 1 à 5 + Devoir maison"],
        "4": ["· Corrigé du devoir maison + Exercices 1 à 5"],
    }

    for line in sommaire_lines:
        line = line.strip()
        if not line:
            continue
        m = re.match(r'^(\d+)\.\s*(.*)', line)
        if m:
            num = m.group(1); label = m.group(2)
            # Titre de section
            p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after  = Pt(2)
            r = p.add_run(f"Section {num} — {label}")
            r.bold = True; r.font.size = Pt(11); r.font.color.rgb = C_BLEU_TITRE
            # Sous-points
            for sub in enriched.get(num, []):
                ps = cell.add_paragraph()
                ps.paragraph_format.space_before = Pt(0)
                ps.paragraph_format.space_after  = Pt(1)
                ps.paragraph_format.left_indent  = Pt(14)
                rs = ps.add_run(sub)
                rs.font.size = Pt(10); rs.font.color.rgb = C_GRIS_CLAIR

    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    doc.add_page_break()


# ── HELPERS ───────────────────────────────────────────────────
def _run(para, text, bold=False, italic=False, color=None, size=10.5):
    r = para.add_run(text)
    r.bold = bold; r.italic = italic
    r.font.size = Pt(size)
    if color:
        r.font.color.rgb = color
    return r


def make_box(bg_color):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.cell(0, 0)
    set_cell_bg(cell, bg_color)
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in [("top","70"),("bottom","70"),("left","120"),("right","120")]:
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"), val); m.set(qn("w:type"), "dxa")
        tcMar.append(m)
    tcPr.append(tcMar)
    return tbl, cell


def space_after_box(pts=4):
    doc.add_paragraph().paragraph_format.space_after = Pt(pts)


def _add_para_in_cell(cell, text, bold=False, italic=False,
                       color=None, size=10.5, sb=0, sa=2,
                       align=WD_ALIGN_PARAGRAPH.LEFT, indent_pt=0):
    p = cell.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(sb)
    p.paragraph_format.space_after  = Pt(sa)
    if indent_pt:
        p.paragraph_format.left_indent = Pt(indent_pt)
    _run(p, text, bold=bold, italic=italic,
         color=color or C_GRIS_TEXTE, size=size)
    return p


# ── BANNIÈRE SECTION ─────────────────────────────────────────
def add_banner(text):
    tbl, cell = make_box(BG_SECTION)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    _run(p, text, bold=True, color=C_BLANC, size=13)
    space_after_box(5)


# ── PHASE HEADER ─────────────────────────────────────────────
def add_phase(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(2)
    # Ligne de soulignement
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), '2')
    bottom.set(qn('w:space'), '1'); bottom.set(qn('w:color'), 'D0D8E8')
    pBdr.append(bottom); pPr.append(pBdr)
    _run(p, text, bold=True, color=C_BLEU_TITRE, size=14)


# ── SOUS-TITRE ───────────────────────────────────────────────
def add_sous_titre(text, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after  = Pt(2)
    _run(p, text, bold=True, color=color or C_OR_TITRE, size=11.5)


# ── TEXTE COURANT ────────────────────────────────────────────
def add_text(text, italic=False, color=None, size=11, sb=0, sa=3,
             align=WD_ALIGN_PARAGRAPH.LEFT, indent_pt=0):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(sb)
    p.paragraph_format.space_after  = Pt(sa)
    if indent_pt:
        p.paragraph_format.left_indent = Pt(indent_pt)
    _run(p, text, italic=italic, color=color or C_GRIS_TEXTE, size=size)


# ── BOÎTE GÉNÉRIQUE ──────────────────────────────────────────
def add_simple_box(lines, bg, title=None, title_color=None,
                   title_size=10, title_icon=""):
    tbl, cell = make_box(bg)
    first = True
    if title:
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(5)
        txt = (title_icon + "  " + title) if title_icon else title
        _run(p, txt, bold=True, color=title_color or C_BLEU_TITRE, size=title_size)
        first = False
    for line in lines:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE HARKNESS enrichie ──────────────────────────────────
def add_harkness_box(title, lines):
    tbl, cell = make_box(BG_HARKNESS)
    # Titre
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(5)
    _run(p, "▶  " + title, bold=True, color=C_BLEU_TITRE, size=12)
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        if line.startswith("Q :") or line.startswith("Q:"):
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
        elif line.startswith("Guide") or line.startswith("Piste"):
            _run(p, line, italic=True, color=C_GRIS_CLAIR, size=10.5)
        elif line.startswith("Corrigé") or line.startswith("Corrige"):
            _run(p, line, bold=True, color=C_VERT_OK, size=10.5)
        else:
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE TEXTE DE DÉCOUVERTE ────────────────────────────────
def add_decouverte_box(eng_lines, trad_lines):
    tbl, cell = make_box(BG_DECOUVERTE)
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(5)
    _run(p0, "📖  Texte de découverte", bold=True, color=C_BLEU_TITRE, size=10)
    for line in eng_lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    if trad_lines:
        p_sep = cell.add_paragraph()
        p_sep.paragraph_format.space_before = Pt(5)
        p_sep.paragraph_format.space_after  = Pt(2)
        _run(p_sep, "— Traduction :", bold=True, color=C_BLEU_TITRE, size=10)
        for line in trad_lines:
            p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after  = Pt(2)
            _run(p, line, italic=True, color=C_GRIS_CLAIR, size=10.5)
    space_after_box()


# ── BOÎTE À RETENIR ──────────────────────────────────────────
def add_retenir_box(title, lines):
    tbl, cell = make_box(BG_RETENIR)
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(5)
    _run(p0, "📘  " + title, bold=True, color=C_VERT_OK, size=10)
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE ATTENTION ──────────────────────────────────────────
def add_attention_box(title, lines):
    tbl, cell = make_box(BG_ATTENTION)
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(5)
    _run(p0, "⚠  " + title, bold=True, color=C_OR_TITRE, size=10)
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        # ✗ en rouge, ✓ en vert
        if '✗' in line or '✓' in line:
            parts = re.split(r'(→)', line)
            for part in parts:
                part = part.strip()
                if not part:
                    continue
                if '✗' in part:
                    _run(p, part + "  ", color=C_ROUGE_ERR, size=10.5)
                elif '✓' in part:
                    _run(p, part, bold=True, color=C_VERT_OK, size=10.5)
                elif part == '→':
                    _run(p, "  →  ", italic=True, color=C_GRIS_CLAIR, size=10.5)
                else:
                    _run(p, part, color=C_GRIS_TEXTE, size=10.5)
        else:
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE DÉFINITION ─────────────────────────────────────────
def add_def_box(title, lines):
    tbl, cell = make_box(BG_DEF)
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(5)
    _run(p0, title, bold=True, color=C_BLEU_TITRE, size=11)
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE TRACE ÉCRITE ───────────────────────────────────────
def add_trace_box(lines):
    tbl, cell = make_box(BG_TRACE)
    first = True
    for line in lines:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        if re.match(r'^(I|II|III|IV|V)\.\s', line):
            _run(p, line, bold=True, color=C_BLEU_TITRE, size=11)
        elif re.match(r'^✎', line):
            _run(p, line, bold=True, color=C_BLEU_TITRE, size=10)
        elif re.match(r'^⚠', line):
            _run(p, line, bold=True, color=C_OR_TITRE, size=10.5)
        else:
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE ANALOGIE ───────────────────────────────────────────
def add_analogie_box(title, lines):
    tbl, cell = make_box(BG_ANALOGIE)
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(5)
    _run(p0, "💡  " + title, bold=True, color=C_ORANGE_ANAL, size=10)
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE SYNTHÈSE ───────────────────────────────────────────
def add_synthese_box(lines):
    tbl, cell = make_box(BG_SYNTHESE)
    first = True
    for line in lines:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        if re.match(r'^5 points|^5 mots|^3 erreurs', line):
            _run(p, line, bold=True, color=C_BLEU_TITRE, size=11)
        elif re.match(r'^\s*\d+\.', line):
            m = re.match(r'^(\s*\d+\.\s*)(.*)', line)
            if m:
                _run(p, m.group(1), bold=True, color=C_BLEU_TITRE, size=10.5)
                _run(p, m.group(2), color=C_GRIS_TEXTE, size=10.5)
        elif re.match(r'^\s*·', line):
            m = re.match(r'^(\s*·\s*)(\S+)(.*)', line)
            if m:
                _run(p, m.group(1), bold=True, color=C_OR_TITRE, size=10.5)
                _run(p, m.group(2), bold=True, color=C_GRIS_TEXTE, size=10.5)
                _run(p, m.group(3), color=C_GRIS_CLAIR, size=10.5)
        elif re.match(r'^\s*✗', line):
            m = re.match(r'^(\s*✗\s*)(.*?)\s*→\s*(✓.*)', line)
            if m:
                _run(p, m.group(1) + m.group(2) + "  ", bold=True, color=C_ROUGE_ERR, size=10.5)
                _run(p, "→  ", italic=True, color=C_GRIS_CLAIR, size=10.5)
                _run(p, m.group(3), bold=True, color=C_VERT_OK, size=10.5)
            else:
                _run(p, line, color=C_ROUGE_ERR, size=10.5)
        else:
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE EXERCICE GUIDÉ ─────────────────────────────────────
def add_exo_guide_box(title, lines):
    tbl, cell = make_box(BG_EXO_GUIDE)
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(5)
    _run(p0, "◉  " + title, bold=True, color=C_BLEU_TITRE, size=12)
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        if re.match(r'^Notion ciblée\s*:', line):
            m = re.match(r'^(Notion ciblée\s*:\s*)(.*)', line)
            if m:
                _run(p, m.group(1), bold=True, color=C_OR_TITRE, size=10.5)
                _run(p, m.group(2), italic=True, color=C_GRIS_TEXTE, size=10.5)
        elif re.match(r'^(Énoncé|MÉTHODE|Démarche)', line):
            _run(p, line, bold=True, color=C_GRIS_TEXTE, size=10.5)
        elif re.match(r'^Étape \d+', line):
            m = re.match(r'^(Étape \d+ · )(.*)', line)
            if m:
                _run(p, m.group(1), bold=True, color=C_BLEU_TITRE, size=10.5)
                _run(p, m.group(2), color=C_GRIS_TEXTE, size=10.5)
            else:
                _run(p, line, color=C_GRIS_TEXTE, size=10.5)
        elif re.match(r'^✔', line):
            m = re.match(r'^(✔[^:]*:\s*)(.*)', line)
            if m:
                _run(p, m.group(1), bold=True, color=C_VERT_OK, size=10.5)
                _run(p, m.group(2), color=C_GRIS_TEXTE, size=10.5)
            else:
                _run(p, line, bold=True, color=C_VERT_OK, size=10.5)
        else:
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE EXERCICE ENTRAÎNEMENT ──────────────────────────────
def add_exo_ent_box(title, lines):
    tbl, cell = make_box(BG_EXO_ENT)
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(5)
    _run(p0, "◎  " + title, bold=True, color=C_BLEU_TITRE, size=12)
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        if re.match(r'^Notions travaillées\s*:', line):
            m = re.match(r'^(Notions travaillées\s*:\s*)(.*)', line)
            if m:
                _run(p, m.group(1), bold=True, color=C_OR_TITRE, size=10.5)
                _run(p, m.group(2), italic=True, color=C_GRIS_TEXTE, size=10.5)
        elif re.match(r'^GUIDE|^Guide de réussite', line):
            _run(p, line, bold=True, color=C_OR_TITRE, size=10.5)
        elif re.match(r'^Étape \d+', line):
            m = re.match(r'^(Étape \d+ · )(.*)', line)
            if m:
                _run(p, m.group(1), bold=True, color=C_BLEU_TITRE, size=10.5)
                _run(p, m.group(2), color=C_GRIS_TEXTE, size=10.5)
            else:
                _run(p, line, color=C_GRIS_TEXTE, size=10.5)
        else:
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE DEVOIR MAISON ──────────────────────────────────────
def add_devoir_box(title, lines):
    tbl, cell = make_box(BG_DEVOIR)
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(5)
    _run(p0, "◈  " + title, bold=True, color=C_OR_TITRE, size=12)
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        if re.match(r'^(CONTEXTE|ÉNONCÉ|ENONCE|STRUCTURE|CONTRAINTES|GUIDE PAS)', line):
            _run(p, line, bold=True, color=C_OR_TITRE, size=10.5)
        elif re.match(r'^Durée', line):
            _run(p, line, italic=True, color=C_GRIS_CLAIR, size=10)
        elif re.match(r'^Étape \d+', line):
            m = re.match(r'^(Étape \d+ · )(.*)', line)
            if m:
                _run(p, m.group(1), bold=True, color=C_OR_TITRE, size=10.5)
                _run(p, m.group(2), color=C_GRIS_TEXTE, size=10.5)
            else:
                _run(p, line, color=C_GRIS_TEXTE, size=10.5)
        elif line.startswith('·'):
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
        elif re.match(r'^\d+\.', line):
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
        else:
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── BOÎTE CORRIGÉ ────────────────────────────────────────────
def add_corrige_box(title, lines):
    tbl, cell = make_box(BG_CORRIGE)
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after  = Pt(5)
    _run(p0, "✔  " + title, bold=True, color=C_VERT_OK, size=12)
    for line in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(2)
        if re.match(r'^✔', line):
            _run(p, line, bold=True, color=C_VERT_OK, size=10.5)
        elif re.match(r'^TRADUCTION|^TITRE PROPOSÉ|^TEXTE PROPOSÉ', line):
            _run(p, line, bold=True, color=C_GRIS_TEXTE, size=10.5)
        else:
            _run(p, line, color=C_GRIS_TEXTE, size=10.5)
    space_after_box()


# ── CITATION FINALE ──────────────────────────────────────────
def add_citation(text_en, text_fr, prof_org):
    doc.add_paragraph().paragraph_format.space_after = Pt(10)
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_before = Pt(4)
    p1.paragraph_format.space_after  = Pt(2)
    _run(p1, f'"{text_en}"', italic=True, color=C_BLEU_TITRE, size=11)
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after  = Pt(2)
    _run(p2, f'({text_fr})', italic=True, color=C_GRIS_CLAIR, size=10)
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_before = Pt(2)
    p3.paragraph_format.space_after  = Pt(0)
    _run(p3, prof_org, color=C_GRIS_CLAIR, size=9)


# ============================================================
#  PARSEUR PRINCIPAL
# ============================================================

lines     = COURS.strip().split('\n')
total     = len(lines)
meta      = {}

# ── Extraction en-tête ════...════ ──────────────────────────
header_lines = []
in_header = False; header_done = False; body_start = 0
for idx, ln in enumerate(lines):
    s = ln.strip()
    if re.match(r'^[═]{4,}', s) and not in_header and not header_done:
        in_header = True; continue
    if re.match(r'^[═]{4,}', s) and in_header:
        in_header = False; header_done = True; body_start = idx + 1; break
    if in_header and s:
        header_lines.append(s)

# ── Méta-lignes ─────────────────────────────────────────────
i = body_start
while i < total:
    s = lines[i].strip()
    m = re.match(r'^(TITRE_LECON|COMPETENCE|PROFESSEUR|ORGANISME)\s*:\s*(.+)', s)
    if m:
        meta[m.group(1)] = m.group(2).strip(); i += 1
    elif not s or s == "SOMMAIRE":
        i += 1
        if s == "SOMMAIRE": break
    else:
        break

titre_lecon = meta.get("TITRE_LECON", "")
competence  = meta.get("COMPETENCE",  "")
professeur  = meta.get("PROFESSEUR",  "")
organisme   = meta.get("ORGANISME",   "")

# Numéro de leçon depuis l'en-tête
lecon_num = "Leçon"
if header_lines:
    m_lec = re.search(r'(Leçon\s*\d+)', header_lines[0])
    if m_lec:
        lecon_num = m_lec.group(1)

# Sommaire
sommaire_lines = []
while i < total:
    s = lines[i].strip()
    if not s or re.match(r'^[─═]{4,}', s): break
    sommaire_lines.append(s); i += 1

# ── En-tête / pied / page de garde / sommaire ───────────────
lecon_info = header_lines[0] if header_lines else "Anglais LV1"
add_header(lecon_info)
add_footer(professeur, organisme)
add_page_garde(titre_lecon, competence, professeur, organisme, lecon_num)
add_sommaire(sommaire_lines, {})

# ── Variables d'état ────────────────────────────────────────
cur_type  = None
cur_title = ""
cur_lines = []
in_tilde  = False
tilde_lines = []
in_syn    = False
syn_lines = []

SOUS_TITRES = re.compile(
    r'^(Objectifs|Prérequis|MÉTHODE|GUIDE|CONTRAINTES|CONTEXTE|'
    r'CONSIGNE|STRUCTURE|GUIDE PAS|Énoncé|Titre\s*:|Compétence\s*:|Professeur\s+\w)')


def flush():
    global cur_type, cur_title, cur_lines
    if not cur_type:
        return
    t, ti, ls = cur_type, cur_title, cur_lines

    if   t == "CAPSULE":   add_simple_box(ls, BG_CAPSULE, title=ti,
                               title_color=C_BLEU_TITRE, title_size=10, title_icon="📌")
    elif t == "ANCRAGE":   add_simple_box(ls, BG_ANCRAGE, title=ti,
                               title_color=C_BLEU_TITRE, title_size=10, title_icon="🌍")
    elif t == "HARKNESS":  add_harkness_box(ti, ls)
    elif t == "RETENIR":   add_retenir_box(ti, ls)
    elif t == "ATTENTION": add_attention_box(ti, ls)
    elif t == "DEF":       add_def_box(ti, ls)
    elif t == "SCHEMA":    add_simple_box(ls, BG_SCHEMA, title=ti,
                               title_color=C_BLEU_TITRE, title_size=10, title_icon="📊")
    elif t == "TRACE":     add_trace_box(ls)
    elif t == "ANALOGIE":  add_analogie_box(ti, ls)
    elif t == "DEVOIR":    add_devoir_box(ti, ls)
    elif t == "EXOGUIDE":  add_exo_guide_box(ti, ls)
    elif t == "EXOENT":    add_exo_ent_box(ti, ls)
    elif t == "CORRIGE":   add_corrige_box(ti, ls)

    cur_type = None; cur_title = ""; cur_lines = []


# ── Citation finale : stockage ───────────────────────────────
citation_en = "Speak even if it's not perfect — silence teaches nothing."
citation_fr = "Parle même si c'est imparfait — le silence n'apprend rien."
citation_sig = ""

# ── Boucle principale ────────────────────────────────────────
while i < total:
    raw = lines[i]; s = raw.strip(); i += 1
    if not s:
        continue

    # Méta déjà traitées
    if re.match(r'^(TITRE_LECON|COMPETENCE|PROFESSEUR|ORGANISME)\s*:', s):
        continue

    # Bloc ~~~ texte de découverte
    if s == "~~~":
        if not in_tilde:
            in_tilde = True; tilde_lines = []
        else:
            in_tilde = False
            flush()
            eng_lines = []; trad_lines = []; in_trad = False
            for tl in tilde_lines:
                if tl == "---TRADUCTION---":
                    in_trad = True; continue
                (trad_lines if in_trad else eng_lines).append(tl)
            add_decouverte_box(eng_lines, trad_lines)
        continue
    if in_tilde:
        tilde_lines.append(s); continue

    # SYNTHESE
    if s == "SYNTHESE_DEBUT":
        flush(); in_syn = True; syn_lines = []; continue
    if s == "SYNTHESE_FIN":
        in_syn = False; add_synthese_box(syn_lines); syn_lines = []; continue
    if in_syn:
        syn_lines.append(s); continue

    # Séparateurs
    if re.match(r'^[═]{4,}', s) or re.match(r'^[─]{4,}', s):
        flush(); continue

    # SECTION
    if re.match(r'^SECTION \d+\s*[—–-]', s):
        flush(); add_banner(s); continue

    # Phase
    if re.match(r'^Phase \d+[a-z]?\s*[·.]', s):
        flush(); add_phase(s); continue

    # [BLOC TITRE]
    m = re.match(r'^\[(.+)\]$', s)
    if m:
        flush(); bl = m.group(1); tu = bl.upper()
        if "CAPSULE" in tu:
            cur_type = "CAPSULE";  cur_title = bl
        elif "ANCRAGE" in tu:
            cur_type = "ANCRAGE";  cur_title = bl
        elif "SCHÉMA" in tu or "SCHEMA" in tu:
            cur_type = "SCHEMA";   cur_title = bl
        else:
            cur_type = "SCHEMA";   cur_title = bl
        continue

    # HARKNESS
    if re.match(r'^[▶►]\s*HARKNESS', s):
        flush()
        cur_type  = "HARKNESS"
        cur_title = s.lstrip("▶► ").strip()
        continue

    # À RETENIR
    if re.match(r'^[┌].*À RETENIR', s):
        flush()
        tc = re.sub(r'^[┌─\s]+', '', s).rstrip('─┐').strip()
        cur_type = "RETENIR"; cur_title = tc; continue
    if re.match(r'^[└]', s):
        flush(); continue

    # ATTENTION
    if re.match(r'^⚠', s):
        flush()
        cur_type  = "ATTENTION"
        cur_title = s.lstrip("⚠ ").strip()
        continue

    # TRACE ÉCRITE
    if re.match(r'^✎', s):
        flush(); cur_type = "TRACE"; cur_lines = [s]; continue

    # ANALOGIE
    if re.match(r'^~\s*ANALOGIE', s):
        if cur_type == "TRACE":
            flush()
        cur_type  = "ANALOGIE"
        cur_title = s.lstrip("~ ").strip()
        cur_lines = []
        continue

    # DÉFINITION DPFC
    if re.match(r'^◆\s*DÉFINITION', s):
        flush(); cur_type = "DEF"; cur_title = s.lstrip("◆ ").strip(); continue

    # EXERCICE GUIDÉ
    if re.match(r'^◉', s):
        flush(); cur_type = "EXOGUIDE"; cur_title = s.lstrip("◉ ").strip(); continue

    # EXERCICE ENTRAÎNEMENT
    if re.match(r'^◎', s):
        flush(); cur_type = "EXOENT"; cur_title = s.lstrip("◎ ").strip(); continue

    # DEVOIR MAISON
    if re.match(r'^◈', s):
        flush(); cur_type = "DEVOIR"; cur_title = s.lstrip("◈ ").strip(); continue

    # CORRIGÉ
    if re.match(r'^✔\s*CORRIGÉ', s):
        flush(); cur_type = "CORRIGE"; cur_title = s.lstrip("✔ ").strip(); continue

    # Citation finale
    if "silence teaches nothing" in s:
        flush()
        # Extraire la citation anglaise
        m_en = re.search(r'"([^"]+)"', s)
        if m_en:
            citation_en = m_en.group(1)
        continue
    if "silence n'apprend rien" in s:
        m_fr = re.search(r'\(([^)]+)\)', s)
        if m_fr:
            citation_fr = m_fr.group(1)
        continue
    # Ligne signature finale
    if re.match(r'^Prof\.?\s+\w+.*DPFC', s):
        citation_sig = s
        continue

    # Contenu d'un bloc ouvert
    if cur_type is not None:
        cur_lines.append(s)
    else:
        # Hors bloc
        if SOUS_TITRES.match(s):
            add_sous_titre(s)
        else:
            add_text(s, color=C_GRIS_TEXTE, size=11)

flush()

# Citation finale
add_citation(citation_en, citation_fr,
             citation_sig or f"{professeur}  ·  {organisme}")

# ============================================================
#  SAUVEGARDE
# ============================================================
nom = "cours_lpa.docx"
if header_lines:
    nom = re.sub(r'[^\w\s·—-]', '', header_lines[0])
    nom = re.sub(r'[\s·—]+', '_', nom).strip('_') + ".docx"

doc.save(nom)
print(f"✅ Document généré : {nom}")
try:
    from google.colab import files
    files.download(nom)
except ImportError:
    print("(Hors Colab — fichier sauvegardé localement)")
