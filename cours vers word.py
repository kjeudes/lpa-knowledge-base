# ============================================================
#  CONVERTISSEUR DE COURS → WORD  (LPA / DPFC Côte d'Ivoire)
#  Google Colab — python-docx  — VERSION 4 (style modèle v2)
#
#  UTILISATION :
#  1. Installe : !pip install python-docx -q
#  2. Colle ton cours dans la variable COURS ci-dessous
#  3. Exécute → téléchargement automatique
#
#  FORMAT DU COURS — BALISES RECONNUES :
#
#  EN-TÊTE (les 2 formats sont acceptés) :
#  ════...════  ligne titre  ════...════     ← en-tête ASCII
#  TITRE_LECON: ...  COMPETENCE: ...        ← méta-lignes optionnelles
#  PROFESSEUR: ...   ORGANISME: ...         ← (auto-extraits de l'en-tête si absents)
#
#  SOMMAIRE (2 formats) :
#  "SOMMAIRE\n  1. Avant la leçon\n  2. La leçon..."   ← ancien format numéroté
#  "SOMMAIRE\n  Section 1 — ...\n    · Phase 1..."      ← nouveau format avec ·
#
#  SECTION 1 — AVANT LA LEÇON
#  ─────────────────────────────────
#  [CAPSULE DE REPRISE — ...]        ← avec ou sans crochets
#   CAPSULE DE REPRISE — ...         ← les deux marchent
#  [ANCRAGE IVOIRIEN]                ← avec ou sans crochets
#   ANCRAGE IVOIRIEN                 ← les deux marchent
#  ▶ HARKNESS N — ...
#
#  SECTION 2 — LA LEÇON
#  Phase N · ...
#  A. TITRE / B. TITRE / C. TITRE    ← sous-sections Phase 2
#  TABLEAU DES STRUCTURES — ...      ← titre de tableau
#  | col | col | col |               ← tableau markdown (auto-converti)
#  ~~~  texte anglais                ← texte de découverte
#  ---TRADUCTION---
#  texte français
#  ~~~
#  ┌─ À RETENIR — ... ─┐ ... └─┘
#  ⚠ ATTENTION — ...
#  FORMATION DU ...                  ← boîte retenir auto
#  ◆ DÉFINITION DPFC — ...
#  ✎ TRACE ÉCRITE
#  ~ ANALOGIE CI 1 / ~ ANALOGIE CI 2
#  5 points essentiels à retenir :   ← déclenche la boîte synthèse
#  SYNTHESE_DEBUT ... SYNTHESE_FIN   ← ou balises explicites
#  ◉ EXERCICE GUIDÉ N — ...
#
#  SECTION 3 — APRÈS LA LEÇON
#  ◎ EXERCICE N — ...
#  ◈ DEVOIR MAISON — ...
#
#  SECTION 4 — CORRIGÉS COMPLETS
#  ✔ CORRIGÉ — ...
#
#  FIN DU COURS :
#  CITATION FINALE                   ← balise optionnelle
#  "Citation en anglais..."
#  (Traduction française)
#  — Auteur
#  Prof. Evelyne ASSI · ... · DPFC/MENET-FP CI · 2025-2026
# ============================================================

COURS = """✅ Anglais LV1 · Leçon 19 — Endemic diseases · en cours de génération...

---

LPA
Lycée Personnel Autonome

Anglais LV1 · Leçon 19 · Classe de 3ème

ENDEMIC DISEASES
Les maladies endémiques

Compétence C7 · Lecture simple — Hygiène et Santé / Health

Professeur Evelyne ASSI · Coefficient 2
DPFC / MENET-FP Côte d'Ivoire · 2025-2026

════════════════════════════════════════════════════════
  ANGLAIS LV1 · Leçon 19 — ENDEMIC DISEASES
  Niveau : 3ème | Série : Programme commun
  Professeur : Evelyne ASSI · Coefficient : 2
  Compétence : C7 · Lecture simple — Hygiène et Santé / Health
  Programme : DPFC/MENET-FP CI 2025-2026
════════════════════════════════════════════════════════

SOMMAIRE
  Section 1 — Avant la leçon
    · Capsule de reprise (Leçon 18 : Tolerance)
    · Ancrage ivoirien
    · Questions Harkness 1, 2 et 3
  Section 2 — La leçon
    · Phase 1 — Présentation & prérequis + Tableau des structures fondamentales
    · Phase 2 — Découverte guidée
    · Phase 3 — Schémas textuels
    · Phase 4 — Définitions DPFC
    · Phase 5 — Trace écrite + Analogies CI
    · Phase 5b — Synthèse
    · Phase 6 — Exercices guidés (2)
  Section 3 — Après la leçon
    · Exercices 1 à 5 + Devoir maison
  Section 4 — Corrigés complets
    · Corrigé du devoir maison + Exercices 1 à 5

────────────────────────────────────────────────────────
SECTION 1 — AVANT LA LEÇON
────────────────────────────────────────────────────────

CAPSULE DE REPRISE — Leçon 18 : Tolerance (La tolérance)

5 points clés à se rappeler :
  1. La tolérance est la capacité d'accepter et de respecter les personnes différentes de nous.
  2. L'intolérance mène aux préjugés, à la discrimination, aux conflits et à la violence.
  3. Les formes d'intolérance : racisme · intolérance religieuse · discrimination ethnique · bullying · sexisme.
  4. Les connecteurs logiques : BECAUSE / BECAUSE OF (cause) · SO / THEREFORE (conséquence) · BUT / HOWEVER / ALTHOUGH (opposition).
  5. Tolérance ne signifie pas tout accepter — on ne tolère pas la violence ni l'injustice.

3 questions flash :
  Q1 : What is the difference between BECAUSE and BECAUSE OF?
       (Quelle est la différence entre BECAUSE et BECAUSE OF ?)
  R1 : BECAUSE + sujet + verbe. BECAUSE OF + nom.
       Exemple : "He left because he was afraid." / "He left because of fear."

  Q2 : Use ALTHOUGH to link these ideas: "We are different. / We can live together."
       (Utilise ALTHOUGH pour relier ces idées.)
  R2 : Although we are different, we can live together.

  Q3 : Name TWO forms of intolerance. (Nomme DEUX formes d'intolérance.)
  R3 : Racism and bullying. / Religious intolerance and sexism. (Autres réponses valides.)

Lien avec la leçon du jour :
  Dans la leçon 18, nous avons vu que la tolérance protège les droits de tous. Parmi ces droits figure le droit à la santé. Aujourd'hui, nous allons apprendre à parler des maladies endémiques en anglais — des maladies qui menacent la santé de millions de personnes en Côte d'Ivoire et en Afrique. C'est le début de la Compétence C7.

────────────────────────────────────────────────────────

ANCRAGE IVOIRIEN

Chaque année en Côte d'Ivoire, des milliers d'enfants tombent malades à cause du paludisme. Dans les quartiers populaires d'Abidjan comme Abobo et Yopougon, les eaux stagnantes autour des maisons deviennent des zones de reproduction pour les moustiques anophèles. À Bouaké, au centre du pays, la bilharziose touche les enfants qui se baignent dans les rivières et les marigots.

Ces maladies ne sont pas des accidents — elles sont endémiques, c'est-à-dire qu'elles sont présentes en permanence dans certaines régions. Les connaître, c'est déjà commencer à les combattre.

Au BEPC, la santé est un thème important. Savoir en parler en anglais est un atout.
Knowing how to talk about health in English is an asset.

Lien DPFC : Ce thème est directement lié aux contenus VIH-SIDA, hygiène et santé publique du programme DPFC CI.

────────────────────────────────────────────────────────

▶ HARKNESS 1 — Qu'est-ce qu'une maladie endémique ?
  Q : What is the difference between an endemic disease, an epidemic and a pandemic?
      (Quelle est la différence entre une maladie endémique, une épidémie et une pandémie ?)
  Guide :
    · Endémique = présente en permanence dans une région précise.
    · Épidémie = augmentation soudaine du nombre de cas dans une région.
    · Pandémie = épidémie qui s'étend à plusieurs pays ou continents.
    · Pense à des exemples : paludisme (endémique en CI), COVID-19 (pandémie).
  Corrigé :
    An endemic disease is a disease that is constantly present in a specific region or population. For example, malaria is endemic in West Africa. An epidemic is a sudden increase in the number of cases of a disease in a particular area. A pandemic is an epidemic that spreads across several countries or continents, such as the COVID-19 pandemic. The key difference is the geographical scale and the regularity of the disease.
    (Une maladie endémique est une maladie constamment présente dans une région ou une population spécifique. Par exemple, le paludisme est endémique en Afrique de l'Ouest. Une épidémie est une augmentation soudaine du nombre de cas d'une maladie dans une zone particulière. Une pandémie est une épidémie qui se propage dans plusieurs pays ou continents, comme la pandémie de COVID-19. La différence principale est l'échelle géographique et la régularité de la maladie.)

▶ HARKNESS 2 — Le paludisme en Côte d'Ivoire
  Q : Why is malaria such a serious problem in Côte d'Ivoire?
      (Pourquoi le paludisme est-il un problème si grave en Côte d'Ivoire ?)
  Guide :
    · Pense aux conditions climatiques de la CI : chaleur, humidité, eaux stagnantes.
    · Pense aux populations les plus touchées : enfants, femmes enceintes.
    · Pense aux conséquences : décès, absentéisme scolaire, pauvreté.
  Corrigé :
    Malaria is a serious problem in Côte d'Ivoire for several reasons. First, the tropical climate — hot and humid — creates ideal conditions for mosquitoes to breed, especially in stagnant water. Second, malaria mainly affects children under five and pregnant women, who are the most vulnerable. Third, it causes many deaths each year and leads to school absenteeism, loss of productivity and increased poverty. Without prevention and treatment, malaria remains one of the biggest health challenges in the country.
    (Le paludisme est un problème grave en Côte d'Ivoire pour plusieurs raisons. Premièrement, le climat tropical — chaud et humide — crée des conditions idéales pour la reproduction des moustiques, notamment dans les eaux stagnantes. Deuxièmement, le paludisme touche principalement les enfants de moins de cinq ans et les femmes enceintes, qui sont les plus vulnérables. Troisièmement, il cause de nombreux décès chaque année et entraîne un absentéisme scolaire, une perte de productivité et une aggravation de la pauvreté. Sans prévention ni traitement, le paludisme reste l'un des plus grands défis sanitaires du pays.)

▶ HARKNESS 3 — Prévention et responsabilité
  Q : What can individuals and communities do to prevent endemic diseases?
      (Que peuvent faire les individus et les communautés pour prévenir les maladies endémiques ?)
  Guide :
    · Pense aux gestes simples : moustiquaires, eau potable, hygiène.
    · Pense au rôle de la communauté : nettoyage des quartiers, sensibilisation.
    · Pense au rôle du gouvernement et des ONG : campagnes de vaccination, distribution de médicaments.
  Corrigé :
    Individuals can protect themselves and their families by sleeping under mosquito nets, drinking clean water, washing their hands regularly and keeping their surroundings clean. Communities can organise neighbourhood clean-up campaigns to eliminate stagnant water and rubbish. Governments and NGOs play a key role by providing free or subsidised treatments, running public awareness campaigns and building health centres in rural areas. Prevention is always better and cheaper than cure.
    (Les individus peuvent se protéger et protéger leurs familles en dormant sous des moustiquaires, en buvant de l'eau potable, en se lavant régulièrement les mains et en maintenant leur environnement propre. Les communautés peuvent organiser des campagnes de nettoyage des quartiers pour éliminer les eaux stagnantes et les ordures. Les gouvernements et les ONG jouent un rôle clé en fournissant des traitements gratuits ou subventionnés, en menant des campagnes de sensibilisation et en construisant des centres de santé dans les zones rurales. La prévention est toujours meilleure et moins coûteuse que le traitement.)

────────────────────────────────────────────────────────
SECTION 2 — LA LEÇON
────────────────────────────────────────────────────────

Phase 1 · Présentation / Prérequis

Titre de la leçon : Endemic Diseases (Les maladies endémiques)
Compétence : C7 — Lecture simple · Hygiène et Santé / Health

Objectifs de la leçon :
  À la fin de cette leçon, tu seras capable de :
  · Nommer et décrire les principales maladies endémiques en Côte d'Ivoire en anglais.
  · Expliquer les causes, les symptômes et les modes de prévention de ces maladies.
  · Utiliser le Present Simple pour décrire des faits généraux sur la santé.
  · Utiliser CAUSE, LEAD TO et PREVENT FROM pour décrire les effets et la prévention des maladies.
  · Lire et comprendre un texte court sur la santé en anglais.

Prérequis nécessaires :
  · Le Present Simple (faits généraux et vérités universelles).
  · Les modaux MUST / SHOULD / HAVE TO (Leçons 16 et 17).
  · Les connecteurs logiques BECAUSE / SO / HOWEVER (Leçon 18).
  · Le vocabulaire du corps humain vu au collège.

────────────────────────────────────────────────────────

TABLEAU DES STRUCTURES FONDAMENTALES — Leçon 19

| Expression / Notion              | Valeur, emploi                                          | Exemple anglais                                                    | Traduction française                                                     |
|---------------------------------|---------------------------------------------------------|--------------------------------------------------------------------|--------------------------------------------------------------------------|
| Endemic disease                 | Maladie endémique                                       | Malaria is an endemic disease in West Africa.                      | Le paludisme est une maladie endémique en Afrique de l'Ouest.            |
| To cause                        | Causer / provoquer                                      | Mosquitoes cause malaria.                                          | Les moustiques causent le paludisme.                                     |
| To lead to                      | Mener à / entraîner                                     | Untreated malaria can lead to death.                               | Le paludisme non traité peut mener à la mort.                            |
| To prevent (from)               | Prévenir / empêcher                                     | Mosquito nets prevent mosquito bites.                              | Les moustiquaires préviennent les piqûres de moustiques.                 |
| Symptom                         | Symptôme                                                | The main symptoms of malaria are fever and headaches.              | Les principaux symptômes du paludisme sont la fièvre et les maux de tête.|
| To be transmitted by            | Être transmis par (Passive Voice)                       | Malaria is transmitted by the Anopheles mosquito.                  | Le paludisme est transmis par le moustique anophèle.                     |
| To be infected by / with        | Être infecté par                                        | Many children are infected with malaria every year.                | De nombreux enfants sont infectés par le paludisme chaque année.         |
| Stagnant water                  | Eau stagnante                                           | Mosquitoes breed in stagnant water.                                | Les moustiques se reproduisent dans les eaux stagnantes.                 |
| Mosquito net                    | Moustiquaire                                            | Sleeping under a mosquito net protects against malaria.            | Dormir sous une moustiquaire protège contre le paludisme.                |
| To be vaccinated against        | Être vacciné contre                                     | Children should be vaccinated against yellow fever.                | Les enfants devraient être vaccinés contre la fièvre jaune.              |
| Clean water / safe water        | Eau potable / eau saine                                 | Drinking clean water prevents many diseases.                       | Boire de l'eau potable prévient de nombreuses maladies.                  |
| Treatment / to treat            | Traitement / traiter                                    | Early treatment can save lives.                                    | Un traitement précoce peut sauver des vies.                              |

────────────────────────────────────────────────────────

Phase 2 · Découverte guidée

A. QU'EST-CE QU'UNE MALADIE ENDÉMIQUE ?
   WHAT IS AN ENDEMIC DISEASE?

An endemic disease is a disease that is regularly and consistently present in a specific region or population.
(Une maladie endémique est une maladie régulièrement et constamment présente dans une région ou une population spécifique.)

It is different from an epidemic (sudden outbreak in one area) or a pandemic (worldwide spread).
(Elle est différente d'une épidémie — augmentation soudaine dans une zone — ou d'une pandémie — propagation mondiale.)

In Côte d'Ivoire and West Africa, several diseases are endemic and affect millions of people every year.
(En Côte d'Ivoire et en Afrique de l'Ouest, plusieurs maladies sont endémiques et touchent des millions de personnes chaque année.)

┌─ À RETENIR ──────────────────────────────────────────┐
  Endemic = endémique (présent en permanence dans une région)
  Epidemic = épidémique (augmentation soudaine dans une zone)
  Pandemic = pandémique (propagation mondiale)
  Disease = maladie
  Illness = maladie / mauvaise santé (terme plus général)
  Infection = infection
└──────────────────────────────────────────────────────┘

────────────────────────────────────────────────────────

B. LES PRINCIPALES MALADIES ENDÉMIQUES EN CÔTE D'IVOIRE
   THE MAIN ENDEMIC DISEASES IN CÔTE D'IVOIRE

1. MALARIA / LE PALUDISME

  What is it? (Qu'est-ce que c'est ?)
    Malaria is a serious infectious disease caused by a parasite called Plasmodium.
    (Le paludisme est une maladie infectieuse grave causée par un parasite appelé Plasmodium.)

  How is it transmitted? (Comment est-il transmis ?)
    It is transmitted by the bite of an infected female Anopheles mosquito.
    (Il est transmis par la piqûre d'un moustique femelle anophèle infecté.)

  Symptoms (Symptômes) :
    · High fever (forte fièvre)
    · Headaches (maux de tête)
    · Chills and shivering (frissons)
    · Vomiting (vomissements)
    · Fatigue (fatigue)

  Who is most affected? (Qui est le plus touché ?)
    Children under five and pregnant women are the most vulnerable.
    (Les enfants de moins de cinq ans et les femmes enceintes sont les plus vulnérables.)

  Prevention (Prévention) :
    · Sleep under a mosquito net. (Dormir sous une moustiquaire.)
    · Eliminate stagnant water around the home. (Éliminer les eaux stagnantes autour de la maison.)
    · Use mosquito repellent. (Utiliser un répulsif anti-moustiques.)
    · Take anti-malaria medication when travelling. (Prendre des médicaments antipaludéens en voyage.)

────────────────────────────────────────────────────────

2. CHOLERA / LE CHOLÉRA

  What is it?
    Cholera is a bacterial infection of the intestines caused by contaminated water or food.
    (Le choléra est une infection bactérienne des intestins causée par de l'eau ou de la nourriture contaminée.)

  How is it transmitted?
    It is transmitted by drinking or eating contaminated water or food.
    (Il est transmis par la consommation d'eau ou de nourriture contaminée.)

  Symptoms :
    · Severe diarrhoea (diarrhée sévère)
    · Vomiting (vomissements)
    · Rapid dehydration (déshydratation rapide)
    · Muscle cramps (crampes musculaires)

  Prevention :
    · Drink only clean, safe water. (Ne boire que de l'eau propre et saine.)
    · Wash hands with soap before eating. (Se laver les mains au savon avant de manger.)
    · Cook food thoroughly. (Bien cuire les aliments.)
    · Avoid eating from street vendors with poor hygiene. (Éviter les vendeurs de rue aux conditions d'hygiène douteuses.)

────────────────────────────────────────────────────────

3. TYPHOID FEVER / LA FIÈVRE TYPHOÏDE

  What is it?
    Typhoid fever is a bacterial infection caused by Salmonella typhi, spread through contaminated food and water.
    (La fièvre typhoïde est une infection bactérienne causée par Salmonella typhi, transmise par des aliments et de l'eau contaminés.)

  Symptoms :
    · Prolonged high fever (fièvre élevée prolongée)
    · Stomach pain (douleurs abdominales)
    · Weakness (faiblesse)
    · Loss of appetite (perte d'appétit)

  Prevention :
    · Drink clean water. (Boire de l'eau propre.)
    · Wash hands regularly. (Se laver régulièrement les mains.)
    · Get vaccinated. (Se faire vacciner.)

────────────────────────────────────────────────────────

4. YELLOW FEVER / LA FIÈVRE JAUNE

  What is it?
    Yellow fever is a viral disease transmitted by the bite of an infected Aedes mosquito.
    (La fièvre jaune est une maladie virale transmise par la piqûre d'un moustique Aedes infecté.)

  Symptoms :
    · Fever (fièvre)
    · Yellowing of the skin and eyes — jaundice (jaunisse : jaunissement de la peau et des yeux)
    · Muscle pain (douleurs musculaires)
    · Bleeding (saignements dans les cas graves)

  Prevention :
    · Vaccination is the best protection. (La vaccination est la meilleure protection.)
    · Use mosquito nets and repellents. (Utiliser des moustiquaires et des répulsifs.)

┌─ À RETENIR ──────────────────────────────────────────┐
  4 maladies endémiques principales en CI :
  1. Malaria / Paludisme → moustique anophèle
  2. Cholera / Choléra → eau / nourriture contaminée
  3. Typhoid fever / Fièvre typhoïde → eau / nourriture contaminée
  4. Yellow fever / Fièvre jaune → moustique Aedes · vaccin disponible
└──────────────────────────────────────────────────────┘

────────────────────────────────────────────────────────

C. GRAMMAIRE — TO CAUSE / TO LEAD TO / TO PREVENT FROM

Ces trois structures sont essentielles pour décrire les maladies, leurs effets et leur prévention.

1. TO CAUSE = causer, provoquer
   Sujet + cause(s) + objet
   · Mosquitoes cause malaria.
     (Les moustiques causent le paludisme.)
   · Contaminated water causes cholera.
     (L'eau contaminée cause le choléra.)
   · Poor hygiene causes many diseases.
     (Une mauvaise hygiène cause de nombreuses maladies.)

2. TO LEAD TO = mener à, entraîner (conséquence grave)
   Sujet + lead(s) to + nom
   · Untreated malaria can lead to death.
     (Le paludisme non traité peut mener à la mort.)
   · Dehydration leads to serious health problems.
     (La déshydratation entraîne de graves problèmes de santé.)
   · Poor nutrition leads to a weak immune system.
     (Une mauvaise nutrition entraîne un système immunitaire affaibli.)

3. TO PREVENT + objet + FROM + verbe-ing = empêcher quelqu'un / quelque chose de faire
   Sujet + prevent(s) + objet + from + verbe-ing
   · Mosquito nets prevent mosquitoes from biting us.
     (Les moustiquaires empêchent les moustiques de nous piquer.)
   · Vaccination prevents children from getting yellow fever.
     (La vaccination empêche les enfants d'attraper la fièvre jaune.)
   · Washing hands prevents germs from spreading.
     (Se laver les mains empêche les germes de se propager.)

⚠ ATTENTION
  TO PREVENT FROM est toujours suivi d'un verbe en -ING.
  Ne jamais écrire : "prevent from bite" → FAUX
  Toujours écrire : "prevent from biting" → CORRECT

⚠ ATTENTION — Faux ami !
  "Intoxicated" en anglais ≠ "Intoxiqué" en français (dans le sens médical courant).
  · Intoxicated (EN) = ivre / sous l'effet de l'alcool ou d'une drogue
  · Intoxiqué (FR) dans le sens courant = poisoned (EN) = empoisonné
  Exemple : "He was intoxicated." = "Il était ivre." (PAS "il était intoxiqué alimentairement")
  Pour parler d'une intoxication alimentaire : "food poisoning" (EN).

────────────────────────────────────────────────────────

Phase 3 · Schémas textuels

[SCHÉMA 1 — Les 4 maladies endémiques : tableau comparatif]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Tableau à 4 lignes et 4 colonnes
  · En-têtes de colonnes (fond bleu foncé, texte blanc) :
      DISEASE · CAUSE · SYMPTOMS · PREVENTION
  · Ligne 1 (fond rouge pâle) : MALARIA
      Cause : Anopheles mosquito
      Symptoms : fever · headaches · chills · vomiting
      Prevention : mosquito net · eliminate stagnant water · repellent
  · Ligne 2 (fond orange pâle) : CHOLERA
      Cause : contaminated water / food
      Symptoms : diarrhoea · vomiting · dehydration
      Prevention : clean water · wash hands · cook food
  · Ligne 3 (fond jaune pâle) : TYPHOID FEVER
      Cause : Salmonella typhi (contaminated food / water)
      Symptoms : prolonged fever · stomach pain · weakness
      Prevention : clean water · wash hands · vaccination
  · Ligne 4 (fond vert pâle) : YELLOW FEVER
      Cause : Aedes mosquito
      Symptoms : fever · jaundice · muscle pain · bleeding
      Prevention : vaccination · mosquito net · repellent
  · Bas du tableau : "Prevention is better than cure." (Mieux vaut prévenir que guérir.)
Note générateur : tableau Word, draw.io ou Canva — 4 lignes × 4 colonnes avec couleurs distinctes
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

[SCHÉMA 2 — Carte mentale : MALARIA]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Nœud central (ovale rouge) : "MALARIA / PALUDISME"
  · Branche 1 (gauche, fond bleu) : "CAUSE"
      → Anopheles mosquito (moustique anophèle femelle)
      → Stagnant water (eaux stagnantes)
      → Tropical climate (climat tropical)
  · Branche 2 (haut, fond orange) : "SYMPTOMS / SYMPTÔMES"
      → High fever (forte fièvre)
      → Headaches (maux de tête)
      → Chills (frissons)
      → Vomiting (vomissements)
      → Fatigue
  · Branche 3 (droite, fond vert) : "PREVENTION / PRÉVENTION"
      → Mosquito net (moustiquaire)
      → Eliminate stagnant water
      → Repellent (répulsif)
      → Anti-malaria medication
  · Branche 4 (bas, fond violet) : "MOST AFFECTED / PLUS VULNÉRABLES"
      → Children under 5 (enfants de moins de 5 ans)
      → Pregnant women (femmes enceintes)
  · Légende : bleu = cause | orange = symptômes | vert = prévention | violet = personnes à risque
Note générateur : MindMeister ou Canva — disposition en étoile à 4 branches
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

────────────────────────────────────────────────────────

Phase 4 · Définitions DPFC

◆ DÉFINITION DPFC — Endemic disease / Maladie endémique
  A disease that is regularly and consistently present in a specific geographical region or population, as opposed to an epidemic (sudden outbreak) or a pandemic (worldwide spread).
  (Une maladie régulièrement et constamment présente dans une région géographique ou une population spécifique, par opposition à une épidémie — augmentation soudaine — ou à une pandémie — propagation mondiale.)

◆ DÉFINITION DPFC — Malaria / Le paludisme
  A serious infectious disease caused by the Plasmodium parasite, transmitted to humans through the bite of an infected female Anopheles mosquito. It is the leading cause of illness and death in sub-Saharan Africa.
  (Une maladie infectieuse grave causée par le parasite Plasmodium, transmise à l'homme par la piqûre d'un moustique femelle anophèle infecté. C'est la principale cause de maladie et de décès en Afrique subsaharienne.)

◆ DÉFINITION DPFC — Cholera / Le choléra
  A bacterial intestinal infection caused by Vibrio cholerae, transmitted through contaminated water or food, characterised by severe diarrhoea and rapid dehydration.
  (Une infection intestinale bactérienne causée par Vibrio cholerae, transmise par de l'eau ou des aliments contaminés, caractérisée par une diarrhée sévère et une déshydratation rapide.)

◆ DÉFINITION DPFC — Prevention / La prévention
  The set of measures taken to avoid the appearance or spread of a disease, including vaccination, hygiene practices, use of protective equipment and health education.
  (L'ensemble des mesures prises pour éviter l'apparition ou la propagation d'une maladie, notamment la vaccination, les pratiques d'hygiène, l'utilisation d'équipements de protection et l'éducation à la santé.)

◆ DÉFINITION DPFC — Symptom / Symptôme
  A physical or mental sign that indicates the presence of a disease or health problem, such as fever, headache or vomiting.
  (Un signe physique ou mental indiquant la présence d'une maladie ou d'un problème de santé, comme la fièvre, les maux de tête ou les vomissements.)

◆ DÉFINITION DPFC — Vaccination / La vaccination
  The introduction of a vaccine into the body to stimulate the immune system and provide protection against a specific disease.
  (L'introduction d'un vaccin dans l'organisme pour stimuler le système immunitaire et fournir une protection contre une maladie spécifique.)

────────────────────────────────────────────────────────

Phase 5 · Trace écrite + Analogies CI

✎ TRACE ÉCRITE (à recopier dans le cahier)

I. VOCABULAIRE CLÉS — ENDEMIC DISEASES / LES MALADIES ENDÉMIQUES

  Endemic disease = maladie endémique
  Epidemic = épidémie
  Pandemic = pandémie
  Disease / illness = maladie
  Infection = infection
  Symptom = symptôme
  Fever = fièvre
  Headache = mal de tête
  Vomiting = vomissements
  Diarrhoea = diarrhée
  Dehydration = déshydratation
  Fatigue = fatigue
  Mosquito = moustique
  Mosquito net = moustiquaire
  Stagnant water = eau stagnante
  Clean water / safe water = eau potable
  Vaccination / vaccine = vaccination / vaccin
  Treatment = traitement
  To prevent = prévenir
  To be transmitted by = être transmis par
  To be infected with = être infecté par
  To cause = causer
  To lead to = mener à / entraîner
  Jaundice = jaunisse
  Repellent = répulsif (anti-moustiques)

II. LES 4 MALADIES ENDÉMIQUES PRINCIPALES EN CI

  1. MALARIA (paludisme) :
     Cause : Anopheles mosquito (moustique anophèle)
     Symptoms : fever · headaches · chills · vomiting · fatigue
     Prevention : mosquito net · eliminate stagnant water · repellent

  2. CHOLERA (choléra) :
     Cause : contaminated water / food (eau / nourriture contaminée)
     Symptoms : severe diarrhoea · vomiting · rapid dehydration
     Prevention : clean water · wash hands · cook food well

  3. TYPHOID FEVER (fièvre typhoïde) :
     Cause : Salmonella typhi — contaminated water / food
     Symptoms : prolonged fever · stomach pain · weakness · loss of appetite
     Prevention : clean water · wash hands · vaccination

  4. YELLOW FEVER (fièvre jaune) :
     Cause : Aedes mosquito (moustique Aedes)
     Symptoms : fever · jaundice · muscle pain · bleeding
     Prevention : vaccination · mosquito net · repellent

III. GRAMMAIRE — TO CAUSE / TO LEAD TO / TO PREVENT FROM

  TO CAUSE = causer :
    Mosquitoes cause malaria.
    (Les moustiques causent le paludisme.)

  TO LEAD TO = mener à / entraîner :
    Untreated malaria can lead to death.
    (Le paludisme non traité peut mener à la mort.)

  TO PREVENT + objet + FROM + verbe-ING = empêcher de :
    Mosquito nets prevent mosquitoes from biting us.
    (Les moustiquaires empêchent les moustiques de nous piquer.)
    ⚠ Toujours verbe en -ING après FROM !

IV. FAUX AMI À RETENIR

  Intoxicated (EN) ≠ Intoxiqué (FR) dans le sens médical courant
  Intoxicated (EN) = ivre / sous l'effet de l'alcool
  Food poisoning (EN) = intoxication alimentaire (FR)
  Exemple : "He had food poisoning after eating at the market."
  = "Il a eu une intoxication alimentaire après avoir mangé au marché."

────────────────────────────────────────────────────────

~ ANALOGIE CI 1
  Les maladies endémiques en Côte d'Ivoire, c'est comme la pluie en saison des pluies à Abidjan : on sait qu'elle va venir, elle revient chaque année, on s'y attend. La différence avec une épidémie, c'est que la pluie saisonnière est prévisible — comme le paludisme. Une épidémie, c'est plutôt comme une crue soudaine qui surprend tout le monde.

~ ANALOGIE CI 2
  TO PREVENT FROM, c'est comme le canari (le filtre à eau traditionnel) dans les maisons ivoiriennes : il empêche les impuretés de passer. "The canary prevents impurities from entering the water" — "Le canari empêche les impuretés d'entrer dans l'eau." C'est exactement la structure TO PREVENT + objet + FROM + verbe-ing.

────────────────────────────────────────────────────────

Phase 5b · Synthèse

5 points essentiels à retenir :
  1. Une maladie endémique est constamment présente dans une région — différente d'une épidémie (soudaine) ou d'une pandémie (mondiale).
  2. Les 4 maladies endémiques principales en CI : malaria · cholera · typhoid fever · yellow fever.
  3. La malaria est la plus meurtrière : transmise par le moustique anophèle, elle touche surtout les enfants de moins de 5 ans et les femmes enceintes.
  4. TO CAUSE (causer) · TO LEAD TO (mener à) · TO PREVENT FROM + verbe-ing (empêcher de) — trois structures indispensables pour parler des maladies.
  5. La prévention est toujours la meilleure arme : eau potable · moustiquaire · hygiène · vaccination.

5 mots clés + définition courte :
  · Endemic disease : maladie constamment présente dans une région précise.
  · Symptom : signe physique indiquant la présence d'une maladie.
  · Mosquito net : moustiquaire — protection contre les piqûres de moustiques.
  · Prevention : ensemble de mesures pour éviter l'apparition d'une maladie.
  · Dehydration : perte excessive d'eau dans l'organisme, conséquence grave du choléra.

3 erreurs fréquentes + correction :
  Erreur 1 : Écrire "prevent from bite" (oubli du -ing après FROM).
  ✔ Correction : "prevent from biting" — après FROM, le verbe prend toujours la forme en -ING.

  Erreur 2 : Confondre "lead to" et "cause".
  ✔ Correction : CAUSE décrit l'origine directe d'une maladie (mosquitoes cause malaria). LEAD TO décrit une conséquence grave qui peut s'ensuivre (malaria can lead to death). Les deux ne sont pas interchangeables.

  Erreur 3 : Traduire "intoxicated" par "intoxiqué alimentairement".
  ✔ Correction : Intoxicated = ivre. Pour une intoxication alimentaire, on dit "food poisoning" en anglais.

────────────────────────────────────────────────────────

Phase 6 · Exercices guidés

◉ EXERCICE GUIDÉ 1 — Cause / Lead to / Prevent from · Notion ciblée : Structures grammaticales sur les maladies

  Énoncé :
  Complete the sentences using TO CAUSE, TO LEAD TO or TO PREVENT FROM + the correct form of the verb.
  Complète les phrases en utilisant TO CAUSE, TO LEAD TO ou TO PREVENT FROM + la forme correcte du verbe.

  1. Poor hygiene ___ many diseases. (cause)
  2. If cholera is not treated quickly, it can ___ death. (lead to)
  3. Vaccination ___ children ___ yellow fever. (prevent / get)
  4. Stagnant water ___ mosquitoes to breed and ___ malaria. (allow / cause)
  5. Sleeping under a mosquito net ___ mosquitoes ___ us. (prevent / bite)

  MÉTHODE
  Étape 1 · Identifie la structure demandée pour chaque phrase.
  Étape 2 · TO CAUSE : sujet + causes + objet (Present Simple, accord du sujet).
  Étape 3 · TO LEAD TO : sujet + leads to / can lead to + nom.
  Étape 4 · TO PREVENT FROM : sujet + prevents + objet + from + verbe-ING.
  Étape 5 · Pour la phrase 4 : deux structures dans la même phrase — relis attentivement.
  ✔ Vérification : Vérifie que les verbes après FROM sont bien en -ING.
  ✔ Conclusion :
    1. Poor hygiene causes many diseases.
       (Une mauvaise hygiène cause de nombreuses maladies.)
    2. If cholera is not treated quickly, it can lead to death.
       (Si le choléra n'est pas traité rapidement, il peut mener à la mort.)
    3. Vaccination prevents children from getting yellow fever.
       (La vaccination empêche les enfants d'attraper la fièvre jaune.)
    4. Stagnant water allows mosquitoes to breed and causes malaria.
       (Les eaux stagnantes permettent aux moustiques de se reproduire et causent le paludisme.)
    5. Sleeping under a mosquito net prevents mosquitoes from biting us.
       (Dormir sous une moustiquaire empêche les moustiques de nous piquer.)
  ✔ Ce que cet exercice mobilise : TO CAUSE · TO LEAD TO · TO PREVENT FROM + verbe-ING · Present Simple · accord sujet-verbe · traduction anglais-français.

────────────────────────────────────────────────────────

◉ EXERCICE GUIDÉ 2 — Reading and identifying diseases · Notion ciblée : Compréhension + vocabulaire des maladies

  Énoncé :
  Read each description and identify the disease. Then write ONE prevention measure in English.
  Lis chaque description et identifie la maladie. Puis écris UNE mesure de prévention en anglais.

  1. "This disease is caused by a parasite. It is transmitted by the bite of a female mosquito. The main symptoms are high fever, headaches and chills. It kills thousands of children in Africa every year."
  2. "This disease is caused by bacteria found in contaminated water. The main symptoms are severe diarrhoea, vomiting and rapid dehydration. It can be deadly if not treated immediately."
  3. "This disease is viral. It is transmitted by a mosquito bite. One of its most recognisable symptoms is yellowing of the skin and eyes. There is an effective vaccine against it."

  MÉTHODE
  Étape 1 · Lis chaque description attentivement et cherche les indices : cause · vecteur · symptômes principaux.
  Étape 2 · Compare avec les maladies étudiées dans ta trace écrite, section II.
  Étape 3 · Identifie la maladie et écris son nom en anglais.
  Étape 4 · Pour la mesure de prévention : choisis une mesure appropriée à la maladie identifiée.
  Étape 5 · Formule ta mesure de prévention en phrase complète en anglais.
  ✔ Vérification : Vérifie que ta mesure de prévention correspond bien à la maladie identifiée.
  ✔ Conclusion :
    1. Disease : Malaria (paludisme).
       Prevention : Sleep under a mosquito net to avoid mosquito bites.
       (Dors sous une moustiquaire pour éviter les piqûres de moustiques.)
    2. Disease : Cholera (choléra).
       Prevention : Always drink clean, safe water and wash your hands with soap before eating.
       (Bois toujours de l'eau propre et saine et lave-toi les mains au savon avant de manger.)
    3. Disease : Yellow fever (fièvre jaune).
       Prevention : Get vaccinated against yellow fever — it is the most effective protection.
       (Fais-toi vacciner contre la fièvre jaune — c'est la protection la plus efficace.)
  ✔ Ce que cet exercice mobilise : compréhension écrite · identification des maladies par leurs caractéristiques · vocabulaire cause-symptômes-prévention · production de phrases en anglais · traduction.

────────────────────────────────────────────────────────
SECTION 3 — APRÈS LA LEÇON
────────────────────────────────────────────────────────

◎ EXERCICE 1 — Vocabulary matching · Notions travaillées : Vocabulaire des maladies endémiques

  Match each English word with its French translation.
  Relie chaque mot anglais à sa traduction française.

  1. Stagnant water            a. Vaccin
  2. Dehydration               b. Jaunisse
  3. Vaccine                   c. Eau stagnante
  4. Jaundice                  d. Moustiquaire
  5. Mosquito net              e. Déshydratation

  GUIDE
  Étape 1 · Relis la trace écrite, section I (Vocabulaire clés).
  Étape 2 · Pour chaque mot anglais, cherche sa traduction dans ta liste de vocabulaire.
  Étape 3 · Relie par une flèche ou écris la lettre correspondante.

────────────────────────────────────────────────────────

◎ EXERCICE 2 — Disease identification · Notions travaillées : Les 4 maladies endémiques + leurs caractéristiques

  Write the name of the correct disease next to each clue.
  Écris le nom de la bonne maladie à côté de chaque indice.

  Diseases to use : Malaria · Cholera · Typhoid fever · Yellow fever

  1. I am caused by contaminated water. My most dangerous symptom is rapid dehydration. → ___
  2. I am transmitted by the Aedes mosquito. I cause yellowing of the skin. There is a vaccine against me. → ___
  3. I am caused by Salmonella typhi. I cause prolonged fever and stomach pain. → ___
  4. I am the most common disease in West Africa. I am transmitted by the Anopheles mosquito. I mainly affect children under five. → ___

  GUIDE
  Étape 1 · Relis la trace écrite, section II (Les 4 maladies endémiques).
  Étape 2 · Pour chaque indice, cherche la maladie dont les caractéristiques correspondent.
  Étape 3 · Écris le nom de la maladie en anglais.

────────────────────────────────────────────────────────

◎ EXERCICE 3 — Cause / Lead to / Prevent from · Notions travaillées : Structures grammaticales

  Fill in the blanks with the correct form of CAUSE, LEAD TO or PREVENT FROM.
  Complète les blancs avec la forme correcte de CAUSE, LEAD TO ou PREVENT FROM.

  1. Drinking contaminated water ___ cholera.
  2. If malaria is not treated, it can ___ serious complications and death.
  3. Washing hands with soap ___ germs ___ spreading.
  4. The Anopheles mosquito ___ malaria in millions of people every year.
  5. Early vaccination ___ children ___ contracting yellow fever.

  GUIDE
  Étape 1 · Identifie la relation exprimée : origine directe → CAUSE / conséquence grave → LEAD TO / empêcher → PREVENT FROM.
  Étape 2 · Pour PREVENT FROM : n'oublie pas le verbe en -ING après FROM.
  Étape 3 · Vérifie l'accord sujet-verbe pour CAUSE et LEAD TO au Present Simple.
  Étape 4 · Traduis chaque phrase complétée en français.

────────────────────────────────────────────────────────

◎ EXERCICE 4 — Prevention sentences · Notions travaillées : Production écrite + MUST / SHOULD / HAVE TO

  Write ONE prevention sentence for each disease using MUST, SHOULD or HAVE TO.
  Écris UNE phrase de prévention pour chaque maladie en utilisant MUST, SHOULD ou HAVE TO.

  1. Malaria
  2. Cholera
  3. Typhoid fever
  4. Yellow fever

  GUIDE
  Étape 1 · Relis la trace écrite, section II (mesures de prévention de chaque maladie).
  Étape 2 · Choisis le bon modal : MUST (obligation forte) · SHOULD (conseil) · HAVE TO (obligation externe).
  Étape 3 · Écris ta phrase : sujet + modal + verbe base + complément.
  Étape 4 · Traduis chaque phrase en français.

────────────────────────────────────────────────────────

◎ EXERCICE 5 — Short reading comprehension · Notions travaillées : Compréhension écrite + vocabulaire santé

  Read the text and answer the questions in English.
  Lis le texte et réponds aux questions en anglais.

  TEXT :
  "Every rainy season in Côte d'Ivoire, the number of malaria cases increases dramatically. This is because the rains create pools of stagnant water around homes, which are ideal breeding grounds for Anopheles mosquitoes. The most vulnerable people are children under five and pregnant women. If malaria is not treated quickly, it can lead to anaemia, brain damage and even death. Fortunately, simple measures can prevent it. Sleeping under a treated mosquito net prevents mosquitoes from biting us during the night. Eliminating stagnant water around the home also reduces the risk significantly."

  Questions :
  1. When do malaria cases increase in Côte d'Ivoire? Why?
  2. Who are the most vulnerable people according to the text?
  3. Find ONE sentence with TO LEAD TO in the text. Write it and translate it.
  4. Find ONE sentence with TO PREVENT FROM in the text. Write it and translate it.
  5. According to the text, what are TWO ways to prevent malaria?

  GUIDE
  Étape 1 · Lis le texte une première fois pour comprendre l'idée générale.
  Étape 2 · Pour les questions 3 et 4 : cherche les expressions "lead to" et "prevent from" dans le texte.
  Étape 3 · Pour la question 1 : cherche WHEN et WHY dans le texte.
  Étape 4 · Réponds en anglais avec des phrases complètes.

────────────────────────────────────────────────────────

◈ DEVOIR MAISON — Fight against malaria in our community
  Durée conseillée : 40 minutes

  Sujet :
  Write a short informative text of 10 to 12 sentences in English.
  (Écris un court texte informatif de 10 à 12 phrases en anglais.)

  Imagine you are a health volunteer in your neighbourhood in Abidjan. You are writing a short leaflet (tract) to inform your community about malaria and how to fight it.
  (Imagine que tu es un volontaire de santé dans ton quartier à Abidjan. Tu rédiges un court tract pour informer ta communauté sur le paludisme et comment le combattre.)

  Your leaflet must include :
  · A title (e.g. "Protect Yourself Against Malaria!")
  · What malaria is (definition + cause)
  · The main symptoms (at least 3)
  · Who is most at risk
  · At least 3 prevention measures
  · A call to action (encourage the community to act)

  Contraintes obligatoires :
  · 10 à 12 phrases minimum.
  · Utiliser TO CAUSE, TO LEAD TO et TO PREVENT FROM + verbe-ING (au moins une fois chacun).
  · Utiliser au moins 2 modaux (MUST / SHOULD / HAVE TO).
  · Utiliser au moins 6 mots du vocabulaire de la leçon.
  · Traduire le tract en français après l'avoir écrit en anglais.

  GUIDE (sans corrigé)
  Étape 1 · Relis la trace écrite : définition du paludisme, symptômes, prévention, vocabulaire.
  Étape 2 · Au brouillon : note la définition, 3 symptômes, les personnes à risque, 3 mesures de prévention.
  Étape 3 · Rédige ton tract avec un titre accrocheur, puis développe point par point.
  Étape 4 · Intègre les structures : "Malaria is caused by... / It can lead to... / Mosquito nets prevent mosquitoes from..."
  Étape 5 · Conclus avec un appel à l'action : "Together, we can fight malaria! / We must act now!"
  Étape 6 · Relis, corrige (accord sujet-verbe, -ING après FROM), puis traduis en français.

────────────────────────────────────────────────────────
SECTION 4 — CORRIGÉS COMPLETS
────────────────────────────────────────────────────────

✔ CORRIGÉ — DEVOIR MAISON : Fight against malaria in our community

Proposition de corrigé (tract modèle en anglais) :

PROTECT YOURSELF AGAINST MALARIA!

Dear neighbours, malaria is one of the most dangerous endemic diseases in Côte d'Ivoire. It is caused by a parasite called Plasmodium, which is transmitted by the bite of an infected female Anopheles mosquito.

The main symptoms of malaria are high fever, severe headaches, chills, vomiting and extreme fatigue. If malaria is not treated quickly, it can lead to anaemia, brain damage and even death.

The most vulnerable people in our community are children under five years old and pregnant women. We must pay special attention to protecting them.

Fortunately, we can fight malaria with simple actions. First, you should sleep under a treated mosquito net every night. Mosquito nets prevent mosquitoes from biting us while we sleep. Second, we must eliminate all stagnant water around our homes — old tyres, broken pots and blocked drains are all breeding grounds for mosquitoes. Third, you should use mosquito repellent in the evening and keep your home clean.

If you or your child has a fever, you must go to the health centre immediately. Early treatment saves lives.

Together, we can protect our community. Let us act now — for our health and the health of our children!

Traduction française complète :

PROTÉGEZ-VOUS CONTRE LE PALUDISME !

Chers voisins, le paludisme est l'une des maladies endémiques les plus dangereuses en Côte d'Ivoire. Il est causé par un parasite appelé Plasmodium, qui est transmis par la piqûre d'un moustique femelle anophèle infecté.

Les principaux symptômes du paludisme sont une forte fièvre, de sévères maux de tête, des frissons, des vomissements et une fatigue extrême. Si le paludisme n'est pas traité rapidement, il peut mener à l'anémie, à des lésions cérébrales et même à la mort.

Les personnes les plus vulnérables dans notre communauté sont les enfants de moins de cinq ans et les femmes enceintes. Nous devons accorder une attention particulière à leur protection.

Heureusement, nous pouvons combattre le paludisme avec des actions simples. Premièrement, vous devriez dormir sous une moustiquaire imprégnée chaque nuit. Les moustiquaires empêchent les moustiques de nous piquer pendant notre sommeil. Deuxièmement, nous devons éliminer toutes les eaux stagnantes autour de nos maisons — les vieux pneus, les pots cassés et les caniveaux bouchés sont autant de zones de reproduction pour les moustiques. Troisièmement, vous devriez utiliser un répulsif anti-moustiques le soir et garder votre maison propre.

Si vous ou votre enfant avez de la fièvre, vous devez aller immédiatement au centre de santé. Un traitement précoce sauve des vies.

Ensemble, nous pouvons protéger notre communauté. Agissons maintenant — pour notre santé et la santé de nos enfants !

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 1 : Vocabulary matching

  1. Stagnant water → c. Eau stagnante
  2. Dehydration → e. Déshydratation
  3. Vaccine → a. Vaccin
  4. Jaundice → b. Jaunisse
  5. Mosquito net → d. Moustiquaire

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 2 : Disease identification

  1. Cholera (choléra) — eau contaminée · déshydratation rapide.
  2. Yellow fever (fièvre jaune) — moustique Aedes · jaunissement de la peau · vaccin disponible.
  3. Typhoid fever (fièvre typhoïde) — Salmonella typhi · fièvre prolongée · douleurs abdominales.
  4. Malaria (paludisme) — moustique anophèle · maladie la plus répandue · enfants de moins de 5 ans.

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 3 : Cause / Lead to / Prevent from

  1. Drinking contaminated water causes cholera.
     (Boire de l'eau contaminée cause le choléra.)

  2. If malaria is not treated, it can lead to serious complications and death.
     (Si le paludisme n'est pas traité, il peut mener à de graves complications et à la mort.)

  3. Washing hands with soap prevents germs from spreading.
     (Se laver les mains au savon empêche les germes de se propager.)

  4. The Anopheles mosquito causes malaria in millions of people every year.
     (Le moustique anophèle cause le paludisme chez des millions de personnes chaque année.)

  5. Early vaccination prevents children from contracting yellow fever.
     (La vaccination précoce empêche les enfants de contracter la fièvre jaune.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 4 : Prevention sentences

  1. Malaria :
     You must sleep under a mosquito net every night to protect yourself from malaria.
     (Tu dois dormir sous une moustiquaire chaque nuit pour te protéger contre le paludisme.)

  2. Cholera :
     We should always drink clean, safe water and wash our hands with soap before eating.
     (Nous devrions toujours boire de l'eau propre et saine et nous laver les mains au savon avant de manger.)

  3. Typhoid fever :
     You have to drink safe water and get vaccinated against typhoid fever.
     (Tu dois boire de l'eau potable et te faire vacciner contre la fièvre typhoïde.)

  4. Yellow fever :
     Every child should be vaccinated against yellow fever — it is the most effective protection.
     (Chaque enfant devrait être vacciné contre la fièvre jaune — c'est la protection la plus efficace.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 5 : Short reading comprehension

  1. Malaria cases increase during the rainy season. This is because the rains create pools of stagnant water, which are ideal breeding grounds for Anopheles mosquitoes.
     (Les cas de paludisme augmentent pendant la saison des pluies. Cela est dû au fait que les pluies créent des flaques d'eau stagnante, qui sont des zones idéales pour la reproduction des moustiques anophèles.)

  2. The most vulnerable people are children under five and pregnant women.
     (Les personnes les plus vulnérables sont les enfants de moins de cinq ans et les femmes enceintes.)

  3. Sentence with TO LEAD TO :
     "If malaria is not treated quickly, it can lead to anaemia, brain damage and even death."
     (Si le paludisme n'est pas traité rapidement, il peut mener à l'anémie, à des lésions cérébrales et même à la mort.)

  4. Sentence with TO PREVENT FROM :
     "Sleeping under a treated mosquito net prevents mosquitoes from biting us during the night."
     (Dormir sous une moustiquaire imprégnée empêche les moustiques de nous piquer pendant la nuit.)

  5. TWO ways to prevent malaria according to the text :
     · Sleeping under a treated mosquito net.
       (Dormir sous une moustiquaire imprégnée.)
     · Eliminating stagnant water around the home.
       (Éliminer les eaux stagnantes autour de la maison.)

════════════════════════════════════════════════════════
  CITATION FINALE
════════════════════════════════════════════════════════

  "The greatest wealth is health."
  (La plus grande richesse, c'est la santé.)

  — Virgil (Virgile)

  Prof. Evelyne ASSI · Anglais LV1 · 3ème · DPFC/MENET-FP CI · 2025-2026
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
    # Grand titre — toujours affiché, même si vide on prend l'en-tête
    titre_affiche = titre if titre else lecon_num
    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4.paragraph_format.space_before = Pt(0)
    p4.paragraph_format.space_after  = Pt(6)
    r4 = p4.add_run(titre_affiche)
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

    num_auto = 0  # compteur si pas de numéro explicite
    for line in sommaire_lines:
        line = line.strip()
        if not line:
            continue
        # Nouveau format : "Section N — ..."
        m2 = re.match(r'^Section\s+(\d+)\s*[—–-]\s*(.*)', line, re.IGNORECASE)
        if m2:
            num = m2.group(1); label = m2.group(2).strip()
            p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after  = Pt(2)
            r = p.add_run(f"Section {num} — {label}")
            r.bold = True; r.font.size = Pt(11); r.font.color.rgb = C_BLEU_TITRE
            for sub in enriched.get(num, []):
                ps = cell.add_paragraph()
                ps.paragraph_format.space_before = Pt(0)
                ps.paragraph_format.space_after  = Pt(1)
                ps.paragraph_format.left_indent  = Pt(14)
                rs = ps.add_run(sub)
                rs.font.size = Pt(10); rs.font.color.rgb = C_GRIS_CLAIR
            continue
        # Sous-point du nouveau format : "    · Phase..."
        if re.match(r'^[·•]\s+', line):
            ps = cell.add_paragraph()
            ps.paragraph_format.space_before = Pt(0)
            ps.paragraph_format.space_after  = Pt(1)
            ps.paragraph_format.left_indent  = Pt(14)
            rs = ps.add_run(line)
            rs.font.size = Pt(10); rs.font.color.rgb = C_GRIS_CLAIR
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
        # Ligne séparatrice visuelle
        p_sep = cell.add_paragraph()
        p_sep.paragraph_format.space_before = Pt(6)
        p_sep.paragraph_format.space_after  = Pt(4)
        pPr = p_sep._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        top_b = OxmlElement('w:top')
        top_b.set(qn('w:val'), 'single'); top_b.set(qn('w:sz'), '2')
        top_b.set(qn('w:space'), '1'); top_b.set(qn('w:color'), 'B0C4D8')
        pBdr.append(top_b); pPr.append(pBdr)
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


# ── TABLEAU STRUCTURES FONDAMENTALES ────────────────────────
def add_tableau_box(title, md_lines):
    """Transforme les lignes markdown | col | col | en vrai tableau Word."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn as _qn

    # Parser les lignes markdown en liste de colonnes
    rows_data = []
    for line in md_lines:
        line = line.strip().strip('|')
        cols = [c.strip() for c in line.split('|')]
        if cols:
            rows_data.append(cols)

    if not rows_data:
        return

    # Titre optionnel au-dessus
    if title:
        p_t = doc.add_paragraph()
        p_t.paragraph_format.space_before = Pt(6)
        p_t.paragraph_format.space_after  = Pt(3)
        _run(p_t, title, bold=True, color=C_BLEU_TITRE, size=11)

    nb_cols = max(len(r) for r in rows_data)
    tbl = doc.add_table(rows=len(rows_data), cols=nb_cols)
    tbl.style = "Table Grid"

    for r_idx, row_data in enumerate(rows_data):
        is_header = (r_idx == 0)
        row = tbl.rows[r_idx]
        for c_idx in range(nb_cols):
            cell = row.cells[c_idx]
            text = row_data[c_idx] if c_idx < len(row_data) else ""
            # Fond
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            shd = OxmlElement("w:shd")
            shd.set(_qn("w:val"), "clear")
            shd.set(_qn("w:color"), "auto")
            shd.set(_qn("w:fill"), "183E82" if is_header else ("F0F4FA" if r_idx % 2 == 0 else "FFFFFF"))
            tcPr.append(shd)
            # Texte
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after  = Pt(2)
            r = p.add_run(text)
            r.bold = is_header
            r.font.size = Pt(9.5)
            r.font.color.rgb = C_BLANC if is_header else C_GRIS_TEXTE

    doc.add_paragraph().paragraph_format.space_after = Pt(5)


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

# Si TITRE_LECON absent : extraire depuis le sommaire (ligne "Section 2")
# ou depuis l'en-tête ════ après le "—"
if not titre_lecon and header_lines:
    # Chercher "Leçon N — TITRE" dans la première ligne d'en-tête
    m_titre = re.search(r'[—–-]\s*(.+)$', header_lines[0])
    if m_titre:
        titre_lecon = m_titre.group(1).strip()

# Si PROFESSEUR absent : chercher dans l'en-tête
if not professeur and header_lines:
    for hl in header_lines:
        m_p = re.search(r'Professeur\s*:\s*(.+)', hl)
        if m_p:
            professeur = m_p.group(1).strip(); break

# Si COMPETENCE absent : chercher dans l'en-tête
if not competence and header_lines:
    for hl in header_lines:
        m_c = re.search(r'Compétence\s*:\s*(.+)', hl)
        if m_c:
            competence = m_c.group(1).strip(); break

# Si ORGANISME absent : chercher dans l'en-tête
if not organisme and header_lines:
    for hl in header_lines:
        m_o = re.search(r'(DPFC/.+)', hl)
        if m_o:
            organisme = m_o.group(1).strip(); break

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
in_trad_libre  = False
trad_libre_lines = []

SOUS_TITRES = re.compile(
    r'^(Objectifs|Prérequis|MÉTHODE|GUIDE|CONTRAINTES|CONTEXTE|'
    r'CONSIGNE|STRUCTURE|GUIDE PAS|Énoncé)')


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
    elif t == "TABLEAU":   add_tableau_box(ti, ls)

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
            # Si la traduction n'était pas dans les ~~~,
            # on active le mode "capture traduction libre"
            add_decouverte_box(eng_lines, trad_lines)
            if not trad_lines:
                in_trad_libre = True
                trad_libre_lines = []
        continue
    if in_tilde:
        tilde_lines.append(s); continue

    # Capture traduction libre après ~~~ sans ---TRADUCTION---
    if in_trad_libre:
        # On arrête à la prochaine boîte ou séparateur
        if (re.match(r'^[┌⚠◆◉◎◈✔✎~▶►\[]', s) or
                re.match(r'^(SECTION|Phase|SYNTHESE)', s) or
                re.match(r'^[─═]{4,}', s)):
            # Injecter les lignes de traduction dans la dernière boîte découverte
            # (on les ajoute via add_text en style traduit)
            in_trad_libre = False
            # Relancer le traitement de cette ligne (ne pas ignorer)
            # on laisse tomber dans la suite du parseur
        else:
            # Ignorer "TRADUCTION COMPLÈTE :" et absorber les lignes
            if not re.match(r'^TRADUCTION', s):
                trad_libre_lines.append(s)
            continue

    # SYNTHESE — avec balises explicites OU détection automatique
    if s == "SYNTHESE_DEBUT" or re.match(r'^5 points essentiels', s):
        flush(); in_syn = True; syn_lines = []
        if s != "SYNTHESE_DEBUT":
            syn_lines.append(s)  # inclure la ligne déclencheuse
        continue
    if s == "SYNTHESE_FIN":
        in_syn = False; add_synthese_box(syn_lines); syn_lines = []; continue
    if in_syn:
        # Fin automatique de synthèse à la prochaine Phase ou Section
        if re.match(r'^(Phase \d|SECTION)', s):
            in_syn = False; add_synthese_box(syn_lines); syn_lines = []
            # Ne pas ignorer cette ligne, la traiter ci-dessous
        else:
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

    # [BLOC TITRE] avec crochets
    m = re.match(r'^\[(.+)\]$', s)
    if m:
        flush(); bl = m.group(1); tu = bl.upper()
        if "CAPSULE" in tu:
            cur_type = "CAPSULE";  cur_title = bl
        elif "ANCRAGE" in tu:
            cur_type = "ANCRAGE";  cur_title = bl
        elif "SCHEMA" in tu or "SCHÉMA" in tu:
            cur_type = "SCHEMA";   cur_title = bl
        else:
            cur_type = "SCHEMA";   cur_title = bl
        continue

    # CAPSULE sans crochets : "CAPSULE DE REPRISE — ..."
    if re.match(r'^CAPSULE\s+DE\s+REPRISE', s):
        flush()
        cur_type = "CAPSULE"; cur_title = s.strip(); cur_lines = []
        continue

    # ANCRAGE sans crochets : "ANCRAGE IVOIRIEN"
    if re.match(r'^ANCRAGE\s+IVOIRIEN', s):
        flush()
        cur_type = "ANCRAGE"; cur_title = "ANCRAGE IVOIRIEN"; cur_lines = []
        continue

    # Titre de tableau juste avant les | (ex: "TABLEAU DES STRUCTURES...")
    if re.match(r'^TABLEAU\s+DES\s+', s):
        flush()
        cur_type = "TABLEAU"; cur_title = s.strip(); cur_lines = []
        continue

    # TABLEAU MARKDOWN — lignes | col | col |
    if re.match(r'^\|', s) and s.count('|') >= 2:
        if cur_type != "TABLEAU":
            flush()
            cur_type = "TABLEAU"; cur_title = ""; cur_lines = []
        # Ignorer lignes séparatrices |---|---|
        if not re.match(r'^\|[-:\s|]+\|', s):
            cur_lines.append(s)
        continue

    # Sous-sections Phase 2 : "A. TITRE", "B. TITRE" etc.
    if re.match(r'^[A-Z]\.\s+[A-ZÀÉÈÊÙÛ\(]', s):
        flush()
        add_sous_titre(s, color=C_BLEU_TITRE)
        continue

    # HARKNESS
    if re.match(r'^[▶►]\s*HARKNESS', s):
        flush()
        cur_type  = "HARKNESS"
        cur_title = s.lstrip("▶► ").strip()
        continue

    # À RETENIR — bordures ┌┐└┘ OU bloc FORMATION ... sans bordures
    if re.match(r'^[┌].*À RETENIR', s):
        flush()
        tc = re.sub(r'^[┌─\s]+', '', s).rstrip('─┐').strip()
        cur_type = "RETENIR"; cur_title = tc; continue
    if re.match(r'^[└]', s):
        flush(); continue
    # Bloc de règles sans bordure (ex: "FORMATION DU PRESENT CONTINUOUS")
    if re.match(r'^FORMATION\s+DU\s+', s):
        flush()
        cur_type  = "RETENIR"
        cur_title = s.strip()
        cur_lines = []
        continue

    # ATTENTION
    if re.match(r'^⚠', s):
        flush()
        cur_type  = "ATTENTION"
        cur_title = s.lstrip("⚠ ").strip()
        continue

    # TRACE ÉCRITE
    if re.match(r'^✎', s):
        flush(); cur_type = "TRACE"; cur_lines = [s]; continue

    # ANALOGIE — flush() systématique (gère 1 ou 2 analogies)
    if re.match(r'^~\s*ANALOGIE', s):
        flush()   # flush TRACE ou ANALOGIE précédente
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

    # Bloc ════ final (pied de cours) → on arrête le contenu
    if re.match(r'^[═]{4,}', s) and i > body_start + 10:
        flush(); continue

    # Balise "CITATION FINALE" → déclenche le mode citation
    if re.match(r'^CITATION\s+FINALE', s):
        flush(); continue

    # Citation entre guillemets (n'importe laquelle, pas forcément la citation LPA)
    if re.match(r'^["""«]', s) and not cur_type:
        flush()
        m_en = re.search(r'["""«]([^"""»]+)["""»]', s)
        if m_en:
            citation_en = m_en.group(1).strip()
        continue
    # Traduction de la citation entre parenthèses
    if re.match(r'^\(', s) and citation_en and not cur_type:
        m_fr = re.search(r'\(([^)]+)\)', s)
        if m_fr:
            citation_fr = m_fr.group(1).strip()
        continue
    # Auteur de la citation "— Prénom Nom"
    if re.match(r'^[—–-]\s+\w', s) and citation_en and not cur_type:
        citation_sig = s.lstrip("—– ").strip()
        continue

    # Mots-clés LPA par défaut
    if "silence teaches nothing" in s:
        flush()
        m_en = re.search(r'"([^"]+)"', s)
        if m_en:
            citation_en = m_en.group(1)
        continue
    if "silence n'apprend rien" in s:
        m_fr = re.search(r'\(([^)]+)\)', s)
        if m_fr:
            citation_fr = m_fr.group(1)
        continue
    # Ligne signature finale Prof. … · DPFC/…
    if re.match(r'^Prof\.?\s+\w+', s) and ('DPFC' in s or 'Anglais' in s or '3ème' in s):
        citation_sig = s
        continue

    # Contenu d'un bloc ouvert
    if cur_type is not None:
        cur_lines.append(s)
    else:
        # Hors bloc
        if SOUS_TITRES.match(s):
            add_sous_titre(s)
        elif re.match(r'^(Titre\s*:|Compétence\s*:|Professeur\s+\w)', s):
            # Lignes de contexte Phase 1 — texte bleu gris léger
            add_text(s, color=C_GRIS_CLAIR, size=11, italic=True)
        elif re.match(r'^LIS CE TEXTE', s):
            # Instruction avant le texte de découverte — texte courant
            add_text(s, color=C_GRIS_TEXTE, size=11, sb=4)
        elif re.match(r'^TRADUCTION', s):
            # Entête "TRADUCTION COMPLÈTE :" hors boîte → ignorer
            pass
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
