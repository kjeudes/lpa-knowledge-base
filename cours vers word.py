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

COURS = """✅ Anglais LV1 · Leçon 20 — Hygiene at school · en cours de génération...

---

LPA
Lycée Personnel Autonome

Anglais LV1 · Leçon 20 · Classe de 3ème

HYGIENE AT SCHOOL
*L'hygiène à l'école*

Compétence C7 · Hygiène et Santé / Health

Professeur Evelyne ASSI · Coefficient 2
DPFC / MENET-FP Côte d'Ivoire · 2025-2026

════════════════════════════════════════════════════════
  ANGLAIS LV1 · Leçon 20 — HYGIENE AT SCHOOL
  Niveau : 3ème | Série : Programme commun
  Professeur : Evelyne ASSI · Coefficient : 2
  Compétence : C7 · Hygiène et Santé / Health
  Programme : DPFC/MENET-FP CI 2025-2026
════════════════════════════════════════════════════════

SOMMAIRE
  Section 1 — Avant la leçon
    · Capsule de reprise (Leçon 19 : Endemic diseases)
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

CAPSULE DE REPRISE — Leçon 19 : Endemic diseases (Maladies endémiques)

5 points clés à retenir :
  1. Une maladie endémique est une maladie présente en permanence dans une région donnée.
  2. En Côte d'Ivoire, le paludisme (malaria) est la maladie endémique la plus répandue.
  3. Vocabulaire clé : disease (maladie) · fever (fièvre) · symptom (symptôme) · prevention (prévention).
  4. On utilise le Present Perfect pour parler d'une maladie que l'on a eue : "I have had malaria twice."
  5. La prévention passe par les moustiquaires, la propreté et la vaccination.

3 questions flash :
  Q1 : What is an endemic disease?
  R1 : A disease that is always present in a specific region. (Une maladie toujours présente dans une région donnée.)

  Q2 : Give two endemic diseases in Côte d'Ivoire.
  R2 : Malaria and typhoid fever. (Le paludisme et la fièvre typhoïde.)

  Q3 : How do we prevent malaria?
  R3 : By sleeping under a mosquito net and keeping our environment clean.
       (En dormant sous une moustiquaire et en gardant notre environnement propre.)

Lien avec la leçon du jour :
  La leçon 19 nous a montré comment certaines maladies se propagent.
  Aujourd'hui, nous allons voir comment l'HYGIÈNE À L'ÉCOLE permet de prévenir
  ces maladies et de maintenir un cadre scolaire sain.

────────────────────────────────────────────────────────

ANCRAGE IVOIRIEN

Kofi est élève en 3ème au Lycée Moderne de Cocody, à Abidjan.
Chaque matin, avant l'entrée en classe, le surveillant général, M. Bamba,
inspecte les élèves : mains propres, ongles coupés, tenue soignée.
La semaine dernière, la directrice, Mme Adjoua Traoré, a organisé une journée
de nettoyage du lycée : les élèves ont balayé les couloirs, vidé les poubelles
et désinfecté les toilettes.
Au réfectoire, la cuisinière, Mariam, porte toujours des gants et un bonnet.

→ Pourquoi toutes ces règles existent-elles ?
→ Que se passerait-il si personne ne les respectait ?
→ C'est exactement ce que nous allons étudier aujourd'hui en anglais.

────────────────────────────────────────────────────────

▶ HARKNESS 1 — Personal hygiene at school
  Q : Why is personal hygiene important at school?
      (Pourquoi l'hygiène personnelle est-elle importante à l'école ?)
  Guide :
    · Pense aux maladies qui se transmettent par contact (touching / le contact).
    · Pense à la concentration en classe quand on est malade.
    · Pense au respect des autres élèves.
  Corrigé :
    Personal hygiene is important at school because students spend many hours
    together in closed spaces. Poor hygiene can spread diseases like flu,
    diarrhoea or skin infections. A clean student is also more comfortable,
    more confident and more focused in class.
    (L'hygiène personnelle est importante à l'école car les élèves passent
    de nombreuses heures ensemble dans des espaces fermés. Une mauvaise hygiène
    peut propager des maladies comme la grippe, la diarrhée ou les infections
    cutanées. Un élève propre est aussi plus à l'aise, plus confiant et plus
    concentré en classe.)

▶ HARKNESS 2 — School environment
  Q : Who is responsible for keeping the school clean?
      (Qui est responsable de la propreté de l'école ?)
  Guide :
    · Pense aux élèves / the students.
    · Pense aux enseignants et à la direction / the teachers and the headmaster.
    · Pense aux agents d'entretien / the cleaning staff.
  Corrigé :
    Keeping the school clean is everyone's responsibility. Students must not
    litter and must use dustbins. Teachers must set a good example. The cleaning
    staff must sweep classrooms and disinfect toilets regularly. The headmaster
    must put hygiene rules in place and make sure they are respected.
    (Garder l'école propre est la responsabilité de tous. Les élèves ne doivent
    pas jeter des ordures et doivent utiliser les poubelles. Les enseignants
    doivent donner le bon exemple. Le personnel d'entretien doit balayer les
    salles et désinfecter les toilettes régulièrement. Le directeur doit mettre
    en place des règles d'hygiène et veiller à leur respect.)

▶ HARKNESS 3 — Hygiene and academic performance
  Q : Can poor hygiene affect a student's results at school?
      (Une mauvaise hygiène peut-elle affecter les résultats d'un élève ?)
  Guide :
    · Pense aux absences pour maladie / absences due to illness.
    · Pense à l'image sociale / social image among classmates.
    · Pense à l'inconfort physique en classe / physical discomfort in class.
  Corrigé :
    Yes, poor hygiene can seriously affect a student's academic performance.
    A student who is often sick misses many classes and falls behind. Bad smells
    or dirty clothes can also cause social exclusion, which leads to stress and
    lack of concentration. A hygienic environment, on the contrary, promotes
    well-being and better learning.
    (Oui, une mauvaise hygiène peut gravement affecter les résultats scolaires
    d'un élève. Un élève souvent malade manque beaucoup de cours et prend du
    retard. Les mauvaises odeurs ou les vêtements sales peuvent aussi entraîner
    une exclusion sociale, qui génère du stress et un manque de concentration.
    Un environnement hygiénique, au contraire, favorise le bien-être et un
    meilleur apprentissage.)

────────────────────────────────────────────────────────
SECTION 2 — LA LEÇON
────────────────────────────────────────────────────────

Phase 1 · Présentation / Prérequis

Titre : Hygiene at school — L'hygiène à l'école
Compétence : C7 · Hygiène et Santé

Objectifs de la leçon :
  → Décrire les règles d'hygiène personnelle et collective à l'école en anglais.
  → Utiliser le modal MUST / MUST NOT pour exprimer l'obligation et l'interdiction.
  → Utiliser SHOULD pour donner des conseils d'hygiène.
  → Enrichir son vocabulaire sur la santé, la propreté et l'environnement scolaire.

Prérequis nécessaires :
  → Vocabulaire de base : school (école) · clean (propre) · dirty (sale) · sick (malade)
  → Verbes d'action simples : wash (laver) · sweep (balayer) · throw (jeter) · clean (nettoyer)
  → Le Present Simple pour les habitudes
  → Notions de la leçon 19 sur les maladies endémiques

┌─ TABLEAU DES STRUCTURES FONDAMENTALES ──────────────────┐
  Expression / Notion      | Valeur, emploi et exemple traduit
  ─────────────────────────|──────────────────────────────────
  must + verbe base        | Obligation forte
                           | "You must wash your hands."
                           | (Tu dois te laver les mains.)
  ─────────────────────────|──────────────────────────────────
  must not + verbe base    | Interdiction absolue
                           | "You must not litter."
                           | (Tu ne dois pas jeter des ordures.)
  ─────────────────────────|──────────────────────────────────
  should + verbe base      | Conseil / recommandation
                           | "You should brush your teeth twice a day."
                           | (Tu devrais te brosser les dents deux fois par jour.)
  ─────────────────────────|──────────────────────────────────
  should not + verbe base  | Déconseiller
                           | "You should not come to school dirty."
                           | (Tu ne devrais pas venir à l'école sale.)
  ─────────────────────────|──────────────────────────────────
  keep + nom + adjectif    | Maintenir dans un état
                           | "Keep your classroom clean."
                           | (Garde ta salle de classe propre.)
  ─────────────────────────|──────────────────────────────────
  It is important to...    | Souligner l'importance
                           | "It is important to use dustbins."
                           | (Il est important d'utiliser les poubelles.)
└──────────────────────────────────────────────────────────┘

────────────────────────────────────────────────────────

Phase 2 · Découverte guidée

TEXTE DE BASE — À lire et comprendre

  "Hygiene at school is very important. Students must wash their hands before
  eating and after using the toilets. They should keep their uniforms clean
  and their nails short. In the classroom, students must not throw rubbish on
  the floor. They should use the dustbin. The school yard must be swept every
  morning. Teachers should remind students of hygiene rules regularly.
  Good hygiene prevents diseases and helps students to concentrate better."

  TRADUCTION FRANÇAISE :
  "L'hygiène à l'école est très importante. Les élèves doivent se laver les mains
  avant de manger et après être allés aux toilettes. Ils devraient garder leurs
  uniformes propres et leurs ongles courts. En classe, les élèves ne doivent pas
  jeter des ordures par terre. Ils devraient utiliser la poubelle. La cour de
  l'école doit être balayée chaque matin. Les enseignants devraient rappeler
  régulièrement les règles d'hygiène aux élèves. Une bonne hygiène prévient
  les maladies et aide les élèves à mieux se concentrer."

────────────────────────────────────────────────────────

┌─ À RETENIR — MUST vs SHOULD ────────────────────────────┐
  MUST → Obligation absolue, règle imposée
  Prononciation : [mʌst]
  Exemple : "Students MUST wash their hands."
            (Les élèves DOIVENT se laver les mains.)

  SHOULD → Conseil, recommandation, ce qui est souhaitable
  Prononciation : [ʃʊd]
  Exemple : "Students SHOULD drink clean water."
            (Les élèves DEVRAIENT boire de l'eau propre.)

  MUST NOT → Interdiction absolue
  Prononciation : [mʌst nɒt]
  Exemple : "Students MUST NOT spit on the floor."
            (Les élèves NE DOIVENT PAS cracher par terre.)

  RÈGLE : Must et should ne prennent JAMAIS de -s à la 3ème personne.
  ❌ He musts wash → ✅ He must wash
  ❌ She shoulds eat → ✅ She should eat
└──────────────────────────────────────────────────────────┘

⚠ ATTENTION — Faux amis
  · "actually" ≠ "actuellement"
    → actually = en réalité / en fait
    → "actuellement" = currently / nowadays
  · "uniform" se prononce [ˈjuːnɪfɔːm] — le "u" se dit "YOU"
  · "hygiène" en français → "hygiene" en anglais [haɪˈdʒiːn]
    Le "h" se prononce, le "g" est doux comme dans "gel"

────────────────────────────────────────────────────────

VOCABULAIRE ESSENTIEL — Hygiene at school

  ANGLAIS              | PRONONCIATION        | TRADUCTION FR
  ─────────────────────|──────────────────────|───────────────────────
  hygiene              | [haɪˈdʒiːn]          | hygiène
  cleanliness          | [ˈklenlɪnəs]         | propreté
  dustbin / bin        | [ˈdʌstbɪn]           | poubelle
  rubbish / waste      | [ˈrʌbɪʃ]             | ordures / déchets
  soap                 | [soʊp]               | savon
  toothbrush           | [ˈtuːθbrʌʃ]          | brosse à dents
  toothpaste           | [ˈtuːθpeɪst]         | dentifrice
  toilet               | [ˈtɔɪlɪt]            | toilettes
  uniform              | [ˈjuːnɪfɔːm]         | uniforme
  nail                 | [neɪl]               | ongle
  sweat                | [swet]               | sueur / transpiration
  disease              | [dɪˈziːz]            | maladie
  infection            | [ɪnˈfekʃən]          | infection
  prevention           | [prɪˈvenʃən]         | prévention
  environment          | [ɪnˈvaɪərənmənt]     | environnement
  sweep                | [swiːp]              | balayer
  disinfect            | [ˌdɪsɪnˈfekt]        | désinfecter
  litter               | [ˈlɪtər]             | jeter des ordures
  canteen / refectory  | [kænˈtiːn]           | cantine / réfectoire
  schoolyard           | [ˈskuːljɑːrd]        | cour d'école

────────────────────────────────────────────────────────

Phase 3 · Schémas textuels

[SCHÉMA 1 — Les règles d'hygiène à l'école / Hygiene rules at school]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Centre (cercle principal, couleur bleue) : HYGIENE AT SCHOOL
  · 4 branches principales rayonnant du centre :
      Branche 1 (verte) : PERSONAL HYGIENE
        → Sous-éléments : wash hands · cut nails · clean uniform · brush teeth
      Branche 2 (orange) : CLASSROOM HYGIENE
        → Sous-éléments : no litter · use dustbin · clean desks · sweep floor
      Branche 3 (rouge) : TOILET HYGIENE
        → Sous-éléments : flush toilet · wash hands after · keep toilets clean
      Branche 4 (violette) : CANTEEN HYGIENE
        → Sous-éléments : wash hands before eating · clean plates · no waste
  · Chaque branche porte son titre en MAJUSCULES
  · Chaque sous-élément est relié à sa branche par une flèche fine
  · Légende : vert = personnel · orange = collectif classe · rouge = sanitaires · violet = alimentation
Note générateur : Canva (Mind Map) ou draw.io
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

[SCHÉMA 2 — Must / Should : tableau de distinction]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Deux colonnes côte à côte :
      Colonne gauche (fond rouge clair) : MUST / MUST NOT
        → Label en haut : "Obligation / Interdiction"
        → Exemples :
           "You must wash your hands." ✔
           "You must not litter." ✘
      Colonne droite (fond vert clair) : SHOULD / SHOULD NOT
        → Label en haut : "Conseil / Déconseiller"
        → Exemples :
           "You should brush your teeth." ✔
           "You should not come dirty." ✘
  · Flèche double au centre séparant les deux colonnes
  · Titre en haut du tableau : MODALS FOR HYGIENE
  · Traductions françaises en italique sous chaque exemple anglais
Note générateur : Canva (Table layout) ou Word simple
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

────────────────────────────────────────────────────────

Phase 4 · Définitions DPFC

◆ DÉFINITION DPFC — Hygiene (Hygiène)
  EN : Hygiene refers to practices and conditions that help maintain health
  and prevent the spread of diseases.
  FR : L'hygiène désigne les pratiques et les conditions qui permettent
  de maintenir la santé et de prévenir la propagation des maladies.

◆ DÉFINITION DPFC — Modal verb (Verbe modal)
  EN : A modal verb is an auxiliary verb used to express obligation, advice,
  permission or possibility. It does not change form with the subject.
  FR : Un verbe modal est un auxiliaire utilisé pour exprimer l'obligation,
  le conseil, la permission ou la possibilité. Il ne change pas de forme
  selon le sujet.

◆ DÉFINITION DPFC — Must (Devoir — obligation)
  EN : "Must" expresses a strong obligation or a strict rule.
  It is followed by the base form of the verb.
  FR : "Must" exprime une obligation forte ou une règle stricte.
  Il est suivi de la base verbale (infinitif sans "to").

◆ DÉFINITION DPFC — Should (Devoir — conseil)
  EN : "Should" expresses advice or a recommendation.
  It suggests what is the right or sensible thing to do.
  FR : "Should" exprime un conseil ou une recommandation.
  Il suggère ce qu'il est juste ou raisonnable de faire.

◆ DÉFINITION DPFC — Prevention (Prévention)
  EN : Prevention is the action of stopping something unpleasant from happening,
  especially illness, by taking precautions in advance.
  FR : La prévention est l'action d'empêcher quelque chose de désagréable
  de se produire, notamment la maladie, en prenant des précautions à l'avance.

────────────────────────────────────────────────────────

Phase 5 · Trace écrite + Analogies CI

✎ TRACE ÉCRITE (à recopier dans le cahier)

  ANGLAIS LV1 · Leçon 20 · 3ème
  HYGIENE AT SCHOOL — L'hygiène à l'école
  Compétence C7 · Hygiène et Santé · Prof. Evelyne ASSI

  I. VOCABULAIRE ESSENTIEL
  ─────────────────────────
  hygiene [haɪˈdʒiːn] = hygiène
  cleanliness = propreté
  dustbin = poubelle
  rubbish = ordures
  soap = savon
  toothbrush = brosse à dents
  nail = ongle
  sweep = balayer
  disinfect = désinfecter
  litter = jeter des ordures
  disease = maladie
  prevention = prévention

  II. LES MODAUX D'OBLIGATION ET DE CONSEIL
  ──────────────────────────────────────────
  MUST + base verbale → Obligation absolue
    Ex : Students must wash their hands.
         (Les élèves doivent se laver les mains.)

  MUST NOT + base verbale → Interdiction absolue
    Ex : You must not litter in the schoolyard.
         (Tu ne dois pas jeter des ordures dans la cour.)

  SHOULD + base verbale → Conseil / recommandation
    Ex : You should brush your teeth twice a day.
         (Tu devrais te brosser les dents deux fois par jour.)

  SHOULD NOT + base verbale → Déconseiller
    Ex : You should not come to school with a dirty uniform.
         (Tu ne devrais pas venir à l'école avec un uniforme sale.)

  RÈGLE : Must et should sont invariables — jamais de -s à la 3ème personne.

  III. LES RÈGLES D'HYGIÈNE À L'ÉCOLE
  ─────────────────────────────────────
  Personal hygiene (hygiène personnelle) :
    · Wash your hands regularly. (Lave-toi les mains régulièrement.)
    · Keep your uniform clean. (Garde ton uniforme propre.)
    · Cut your nails short. (Coupe tes ongles courts.)
    · Brush your teeth morning and evening. (Brosse-toi les dents matin et soir.)

  Classroom hygiene (hygiène en classe) :
    · Use the dustbin. (Utilise la poubelle.)
    · Do not write on the walls. (N'écris pas sur les murs.)
    · Keep your desk clean. (Garde ton bureau propre.)

  School environment (environnement scolaire) :
    · Sweep the schoolyard every morning. (Balayez la cour chaque matin.)
    · Disinfect the toilets regularly. (Désinfectez les toilettes régulièrement.)
    · Do not litter. (Ne jetez pas d'ordures par terre.)

~ ANALOGIE CI
  Pense aux journées de salubrité organisées dans les quartiers d'Abidjan
  (Yopougon, Abobo, Adjamé) chaque premier samedi du mois.
  Tout le monde sort balayer devant sa maison.
  À l'école, c'est pareil : chaque élève est responsable de son espace.
  "Keep your school clean like you keep your home clean."
  (Garde ton école propre comme tu gardes ta maison propre.)

~ ANALOGIE CI 2
  Au marché de Cocody ou de Treichville, les vendeurs de nourriture
  portent des gants et couvrent leurs plats.
  Pourquoi ? Pour éviter les contaminations.
  À la cantine de l'école, les mêmes règles s'appliquent.
  "Hygiene is not a luxury — it is a duty."
  (L'hygiène n'est pas un luxe — c'est un devoir.)

────────────────────────────────────────────────────────

Phase 5b · Synthèse

5 points essentiels à retenir :
  1. L'hygiène à l'école concerne TOUS les élèves, les enseignants et le personnel.
  2. MUST exprime une obligation ou une interdiction (must not) stricte.
  3. SHOULD exprime un conseil ou une recommandation (should not = déconseiller).
  4. Must et should sont invariables : jamais de -s à la 3ème personne du singulier.
  5. Une bonne hygiène prévient les maladies et améliore les résultats scolaires.

5 mots clés + définition courte :
  · Hygiene → ensemble des pratiques pour rester en bonne santé
  · Must → modal d'obligation absolue (pas de forme en -s)
  · Should → modal de conseil (pas de forme en -s)
  · Prevention → action d'empêcher une maladie avant qu'elle n'arrive
  · Cleanliness → état de ce qui est propre ; qualité d'être propre

3 erreurs fréquentes + correction :
  Erreur 1 : "He musts wash his hands."
  Correction : "He must wash his hands." — must est invariable, jamais de -s.

  Erreur 2 : "You should to brush your teeth."
  Correction : "You should brush your teeth." — après un modal, PAS de "to".

  Erreur 3 : Confondre must (obligation) et should (conseil).
  Correction : "You MUST wash your hands before eating." (règle de santé = obligatoire)
               "You SHOULD drink more water." (conseil = recommandé mais pas imposé)

────────────────────────────────────────────────────────

Phase 6 · Exercices guidés

◉ EXERCICE GUIDÉ 1 — Obligation ou conseil ? · Notion ciblée : Must / Should

  Énoncé :
  Complète les phrases suivantes avec MUST, MUST NOT ou SHOULD.
  Traduis chaque phrase en français.

  a) Students _______ wash their hands before eating.
  b) You _______ litter in the classroom.
  c) Kofi _______ brush his teeth every morning — it is a good habit.
  d) The cleaning staff _______ sweep the schoolyard every day — it is a school rule.
  e) You _______ not come to school with a dirty uniform — it gives a bad image.

  MÉTHODE
  Étape 1 · Lis la phrase et repère le sens : est-ce une règle absolue, une interdiction
            ou simplement un conseil ?
            → Règle / obligation imposée : utilise MUST
            → Interdiction : utilise MUST NOT
            → Conseil / recommandation : utilise SHOULD ou SHOULD NOT
  Étape 2 · Vérifie que le verbe qui suit est bien à la BASE VERBALE (infinitif sans "to").
  Étape 3 · Traduis la phrase complète en français.
  ✔ Vérification : Must et should ne changent pas à la 3ème personne — pas de -s.
  ✔ Conclusion :
    a) Students MUST wash their hands before eating.
       (Les élèves DOIVENT se laver les mains avant de manger.)
    b) You MUST NOT litter in the classroom.
       (Tu NE DOIS PAS jeter des ordures en classe.)
    c) Kofi SHOULD brush his teeth every morning — it is a good habit.
       (Kofi DEVRAIT se brosser les dents chaque matin — c'est une bonne habitude.)
    d) The cleaning staff MUST sweep the schoolyard every day — it is a school rule.
       (Le personnel d'entretien DOIT balayer la cour chaque jour — c'est une règle.)
    e) You SHOULD NOT come to school with a dirty uniform — it gives a bad image.
       (Tu NE DEVRAIS PAS venir à l'école avec un uniforme sale — ça donne une mauvaise image.)
  ✔ Ce que cet exercice mobilise :
    · Distinction must / must not / should / should not
    · Règle d'invariabilité des modaux
    · Traduction anglais → français
    · Vocabulaire de l'hygiène scolaire

────────────────────────────────────────────────────────

◉ EXERCICE GUIDÉ 2 — Rédiger des règles d'hygiène · Notion ciblée : Production écrite + Modaux

  Énoncé :
  Amani est délégué de classe au Lycée Moderne d'Adjamé, à Abidjan.
  Il doit rédiger un panneau d'affichage avec 5 règles d'hygiène pour sa salle de classe.
  Aide-le à écrire ces 5 règles en anglais, en utilisant MUST, MUST NOT et SHOULD.
  Traduis chaque règle en français.

  MÉTHODE
  Étape 1 · Pense aux différents aspects de l'hygiène en classe :
            mains, ordures, bureau, murs, uniforme, nourriture.
  Étape 2 · Pour chaque règle, choisis le modal approprié :
            → Règle stricte : MUST
            → Interdiction : MUST NOT
            → Conseil : SHOULD
  Étape 3 · Construis chaque phrase : Sujet + Modal + Base verbale + Complément.
  Étape 4 · Traduis chaque règle en français.
  ✔ Vérification : Vérifie que tu as utilisé au moins une fois MUST, MUST NOT et SHOULD.
  ✔ Conclusion — Exemple de panneau d'Amani :
    HYGIENE RULES FOR OUR CLASSROOM
    ─────────────────────────────────
    Rule 1 : You must wash your hands before entering the classroom.
             (Tu dois te laver les mains avant d'entrer en classe.)
    Rule 2 : You must not throw rubbish on the floor.
             (Tu ne dois pas jeter des ordures par terre.)
    Rule 3 : You should keep your desk clean and tidy.
             (Tu devrais garder ton bureau propre et rangé.)
    Rule 4 : You must not write on the walls or the desks.
             (Tu ne dois pas écrire sur les murs ou les bureaux.)
    Rule 5 : You should put your waste in the dustbin.
             (Tu devrais mettre tes déchets dans la poubelle.)
  ✔ Ce que cet exercice mobilise :
    · Production écrite guidée en anglais
    · Utilisation des modaux MUST / MUST NOT / SHOULD dans un contexte réel
    · Vocabulaire de l'hygiène et de l'environnement scolaire
    · Traduction anglais → français
    · Ancrage ivoirien (lycée d'Adjamé)

────────────────────────────────────────────────────────
SECTION 3 — APRÈS LA LEÇON
────────────────────────────────────────────────────────

◎ EXERCICE 1 — Vocabulaire de l'hygiène · Notions travaillées : Lexique, traduction

  Relie chaque mot anglais à sa traduction française.
  (Écris le numéro du mot anglais devant la bonne traduction.)

  1. soap                    A. poubelle
  2. dustbin                 B. balayer
  3. toothbrush              C. savon
  4. sweep                   D. désinfecter
  5. disinfect               E. brosse à dents

  GUIDE
  Étape 1 · Lis chaque mot anglais à voix haute en te souvenant de la prononciation vue en classe.
  Étape 2 · Associe chaque mot à la traduction française qui correspond.
  Étape 3 · Vérifie tes réponses en relisant la trace écrite.

────────────────────────────────────────────────────────

◎ EXERCICE 2 — Must / Must not / Should · Notions travaillées : Modaux d'obligation et de conseil

  Complète les phrases avec MUST, MUST NOT ou SHOULD.
  Justifie ton choix entre parenthèses : (obligation) / (interdiction) / (conseil).

  a) You _______ drink clean water every day. It is essential for your health.
  b) Students _______ spit on the classroom floor. It is strictly forbidden.
  c) Awa _______ change her school uniform more often — it would be better.
  d) The teacher _______ remind students to wash their hands. It is part of his duties.
  e) You _______ eat food that has fallen on the dirty floor.

  GUIDE
  Étape 1 · Identifie si la phrase exprime une obligation, une interdiction ou un conseil.
  Étape 2 · Choisis le modal correspondant.
  Étape 3 · Vérifie que le verbe suivant est à la base verbale (pas de -s, pas de "to").

────────────────────────────────────────────────────────

◎ EXERCICE 3 — Questions et réponses · Notions travaillées : Compréhension écrite, structures interrogatives

  Lis ce texte, puis réponds aux questions en anglais.
  (Traduis chaque réponse en français.)

  TEXTE :
  "At Lycée Moderne de Yopougon, students follow strict hygiene rules.
  Every morning, they sweep the classroom and put the waste in the dustbin.
  Before lunch, all students must wash their hands with soap.
  The school director has put posters about hygiene on all the walls.
  Students who do not respect the rules are sent to the headmaster's office."

  Questions :
  Q1 : What do students do every morning?
  Q2 : When must students wash their hands?
  Q3 : What has the school director done?
  Q4 : What happens to students who do not respect hygiene rules?

  GUIDE
  Étape 1 · Lis le texte une première fois pour en comprendre le sens général.
  Étape 2 · Lis chaque question et repère dans le texte la partie qui y répond.
  Étape 3 · Formule ta réponse en anglais avec une phrase complète.
  Étape 4 · Traduis ta réponse en français.

────────────────────────────────────────────────────────

◎ EXERCICE 4 — Correction de phrases erronées · Notions travaillées : Grammaire des modaux, erreurs fréquentes

  Ces phrases contiennent des erreurs. Trouve et corrige chaque erreur.
  Traduis la phrase corrigée en français.

  a) She musts wash her hands before eating.
  b) You should to use the dustbin.
  c) Students must not to litter in the schoolyard.
  d) He don't must come to school dirty.
  e) We should keeps our classroom clean.

  GUIDE
  Étape 1 · Rappelle-toi la règle : après un modal (must / should), on utilise la BASE VERBALE.
  Étape 2 · Rappelle-toi que must et should sont INVARIABLES — jamais de -s ni de conjugaison.
  Étape 3 · Corrige la phrase et traduis-la.

────────────────────────────────────────────────────────

◎ EXERCICE 5 — Production écrite · Notions travaillées : Expression écrite, modaux, vocabulaire hygiène

  Ton école va organiser une Journée de l'Hygiène.
  Le directeur te demande d'écrire un court discours (8 à 10 phrases) pour tes camarades.
  Tu dois :
    · Expliquer pourquoi l'hygiène est importante à l'école (2 phrases)
    · Donner au moins 3 règles d'obligation (MUST / MUST NOT)
    · Donner au moins 2 conseils (SHOULD / SHOULD NOT)
    · Terminer par une phrase d'encouragement

  GUIDE
  Étape 1 · Commence par une phrase d'introduction : "Dear classmates, today..."
  Étape 2 · Explique l'importance de l'hygiène (maladies, concentration, respect).
  Étape 3 · Formule tes règles avec MUST et MUST NOT.
  Étape 4 · Ajoute tes conseils avec SHOULD et SHOULD NOT.
  Étape 5 · Termine par une phrase motivante.

────────────────────────────────────────────────────────

◈ DEVOIR MAISON — A letter about hygiene at school
  Durée conseillée : 30 minutes

  Énoncé :
  Ton ami(e) Moussa, qui habite à Bouaké, t'écrit une lettre pour te dire
  que son école est très sale et que personne ne respecte les règles d'hygiène.
  Tu lui réponds en anglais (12 à 15 phrases) pour :
    · Lui dire que c'est un problème sérieux (2 phrases)
    · Lui donner 4 règles d'hygiène qu'il/elle devrait appliquer (avec MUST / MUST NOT)
    · Lui donner 3 conseils supplémentaires (avec SHOULD / SHOULD NOT)
    · L'encourager à parler au directeur de son école
    · Terminer la lettre de façon amicale

  Contraintes obligatoires :
    ✔ Utiliser MUST au moins 2 fois
    ✔ Utiliser MUST NOT au moins 2 fois
    ✔ Utiliser SHOULD au moins 2 fois
    ✔ Utiliser au moins 8 mots de vocabulaire vus en classe
    ✔ Traduire la lettre complète en français
    ✔ Format lettre : lieu + date / salutation / corps / formule de politesse

  GUIDE
  Étape 1 · Commence par l'en-tête : Abidjan, [date] / Dear Moussa,
  Étape 2 · Réponds à son problème : "I understand your situation. It is a serious problem because..."
  Étape 3 · Donne tes règles avec MUST et MUST NOT (une par phrase).
  Étape 4 · Donne tes conseils avec SHOULD et SHOULD NOT.
  Étape 5 · Encourage-le à parler au directeur : "You should talk to your headmaster..."
  Étape 6 · Termine : "I hope things will get better soon. Your friend, [ton prénom]"
  Étape 7 · Traduis toute la lettre en français.

────────────────────────────────────────────────────────
SECTION 4 — CORRIGÉS COMPLETS
────────────────────────────────────────────────────────

✔ CORRIGÉ — DEVOIR MAISON

  Abidjan, 14th July 2026

  Dear Moussa,

  Thank you for your letter. I understand your situation. It is a serious problem
  because a dirty school can spread diseases and make students sick.

  Here are some hygiene rules that your school must follow :
  Rule 1 : Students must wash their hands with soap before eating and after
           using the toilets.
           (Les élèves doivent se laver les mains avec du savon avant de manger
           et après être allés aux toilettes.)
  Rule 2 : Students must not litter in the classrooms or in the schoolyard.
           (Les élèves ne doivent pas jeter des ordures dans les salles
           ou dans la cour.)
  Rule 3 : The cleaning staff must sweep the classrooms and disinfect the
           toilets every day.
           (Le personnel d'entretien doit balayer les salles et désinfecter
           les toilettes chaque jour.)
  Rule 4 : Students must not eat outside the canteen to avoid leaving waste
           everywhere.
           (Les élèves ne doivent pas manger en dehors de la cantine pour
           éviter de laisser des déchets partout.)

  Here are some additional pieces of advice :
  · You should put dustbins in every classroom and in the schoolyard.
    (Tu devrais mettre des poubelles dans chaque salle et dans la cour.)
  · You should not throw rubbish on the floor even when there is no dustbin nearby.
    (Tu ne devrais pas jeter des ordures par terre même quand il n'y a pas
    de poubelle à proximité.)
  · Students should organise a weekly cleaning day to keep the school environment clean.
    (Les élèves devraient organiser une journée de nettoyage hebdomadaire
    pour maintenir l'environnement scolaire propre.)

  You should talk to your headmaster about this problem. Explain that hygiene
  is important for everyone's health and that the school needs clearer rules.
  (Tu devrais parler de ce problème à ton directeur. Explique que l'hygiène
  est importante pour la santé de tous et que l'école a besoin de règles plus claires.)

  I hope things will get better soon. Stay clean and stay healthy!
  (J'espère que les choses vont bientôt s'améliorer. Reste propre et reste en bonne santé !)

  Your friend,
  Kofi

  ─────────────────────────────────────────────────────
  TRADUCTION COMPLÈTE DE LA LETTRE EN FRANÇAIS :

  Abidjan, le 14 juillet 2026

  Cher Moussa,

  Merci pour ta lettre. Je comprends ta situation. C'est un problème sérieux
  car une école sale peut propager des maladies et rendre les élèves malades.

  Voici quelques règles d'hygiène que ton école doit suivre :
  Règle 1 : Les élèves doivent se laver les mains avec du savon avant de manger
            et après être allés aux toilettes.
  Règle 2 : Les élèves ne doivent pas jeter des ordures dans les salles ou dans la cour.
  Règle 3 : Le personnel d'entretien doit balayer les salles et désinfecter
            les toilettes chaque jour.
  Règle 4 : Les élèves ne doivent pas manger en dehors de la cantine pour
            éviter de laisser des déchets partout.

  Voici quelques conseils supplémentaires :
  · Tu devrais mettre des poubelles dans chaque salle et dans la cour.
  · Tu ne devrais pas jeter des ordures par terre même quand il n'y a pas
    de poubelle à proximité.
  · Les élèves devraient organiser une journée de nettoyage hebdomadaire
    pour maintenir l'environnement scolaire propre.

  Tu devrais parler de ce problème à ton directeur. Explique que l'hygiène
  est importante pour la santé de tous et que l'école a besoin de règles plus claires.

  J'espère que les choses vont bientôt s'améliorer. Reste propre et reste en bonne santé !

  Ton ami,
  Kofi

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 1 (Vocabulaire)

  1 — C : soap = savon
  2 — A : dustbin = poubelle
  3 — E : toothbrush = brosse à dents
  4 — B : sweep = balayer
  5 — D : disinfect = désinfecter

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 2 (Must / Must not / Should)

  a) You MUST drink clean water every day. (obligation — essentiel pour la santé)
     (Tu DOIS boire de l'eau propre chaque jour.)

  b) Students MUST NOT spit on the classroom floor. (interdiction stricte)
     (Les élèves NE DOIVENT PAS cracher par terre en classe.)

  c) Awa SHOULD change her school uniform more often. (conseil)
     (Awa DEVRAIT changer son uniforme scolaire plus souvent.)

  d) The teacher MUST remind students to wash their hands. (obligation — part of his duties)
     (L'enseignant DOIT rappeler aux élèves de se laver les mains.)

  e) You MUST NOT eat food that has fallen on the dirty floor. (interdiction — danger de santé)
     (Tu NE DOIS PAS manger de la nourriture tombée par terre.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 3 (Questions et réponses)

  Q1 : What do students do every morning?
  R1 : Every morning, they sweep the classroom and put the waste in the dustbin.
       (Chaque matin, ils balaient la salle de classe et mettent les déchets dans la poubelle.)

  Q2 : When must students wash their hands?
  R2 : Students must wash their hands before lunch.
       (Les élèves doivent se laver les mains avant le déjeuner.)

  Q3 : What has the school director done?
  R3 : The school director has put posters about hygiene on all the walls.
       (Le directeur de l'école a mis des affiches sur l'hygiène sur tous les murs.)

  Q4 : What happens to students who do not respect hygiene rules?
  R4 : Students who do not respect the rules are sent to the headmaster's office.
       (Les élèves qui ne respectent pas les règles sont envoyés dans le bureau du directeur.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 4 (Correction de phrases erronées)

  a) ❌ She musts wash her hands before eating.
     ✅ She must wash her hands before eating.
        (Elle doit se laver les mains avant de manger.)
        → Erreur : must est invariable, jamais de -s.

  b) ❌ You should to use the dustbin.
     ✅ You should use the dustbin.
        (Tu devrais utiliser la poubelle.)
        → Erreur : après un modal, pas de "to" — base verbale directe.

  c) ❌ Students must not to litter in the schoolyard.
     ✅ Students must not litter in the schoolyard.
        (Les élèves ne doivent pas jeter des ordures dans la cour.)
        → Erreur : après must not, pas de "to".

  d) ❌ He don't must come to school dirty.
     ✅ He must not come to school dirty.
        (Il ne doit pas venir à l'école sale.)
        → Erreur : la négation de must est must not — on n'utilise pas "don't/doesn't".

  e) ❌ We should keeps our classroom clean.
     ✅ We should keep our classroom clean.
        (Nous devrions garder notre salle de classe propre.)
        → Erreur : après should, la base verbale — jamais de -s.

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 5 (Production écrite)

  Exemple de discours pour la Journée de l'Hygiène :

  "Dear classmates,
  Today is our school's Hygiene Day and I am happy to speak to you about
  this important topic.
  (Chers camarades, aujourd'hui est la Journée de l'Hygiène de notre école
  et je suis heureux de vous parler de ce sujet important.)

  Hygiene at school is very important because it prevents the spread of diseases.
  A clean school is also a place where we can concentrate and learn better.
  (L'hygiène à l'école est très importante car elle prévient la propagation
  des maladies. Une école propre est aussi un endroit où nous pouvons mieux
  nous concentrer et apprendre.)

  Here are the rules we must all follow :
  You must wash your hands with soap before eating and after using the toilets.
  You must not litter — always use the dustbin.
  You must not spit on the floor or write on the walls.
  (Voici les règles que nous devons tous respecter :
  Tu dois te laver les mains avec du savon avant de manger et après les toilettes.
  Tu ne dois pas jeter des ordures — utilise toujours la poubelle.
  Tu ne dois pas cracher par terre ou écrire sur les murs.)

  Here is my advice :
  You should keep your uniform clean and your nails short.
  You should not come to school without brushing your teeth.
  (Voici mes conseils :
  Tu devrais garder ton uniforme propre et tes ongles courts.
  Tu ne devrais pas venir à l'école sans t'être brossé les dents.)

  Together, we can make our school a clean and healthy place for everyone.
  Let us all make the effort — starting today!
  (Ensemble, nous pouvons faire de notre école un endroit propre et sain pour tous.
  Faisons tous cet effort — à partir d'aujourd'hui !)

  Thank you."
  (Merci.)

════════════════════════════════════════════════════════
  CITATION FINALE
════════════════════════════════════════════════════════

  "Cleanliness is not next to godliness — it is next to health."
  (La propreté n'est pas proche de la sainteté — elle est proche de la santé.)

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
