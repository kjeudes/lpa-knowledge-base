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

COURS = """✅ Anglais LV1 · Leçon 16 — My rights · en cours de génération...

---

LPA
Lycée Personnel Autonome

Anglais LV1 · Leçon 16 · Classe de 3ème

MY RIGHTS
Mes droits

Compétence C6 · Écoute élaborée — Droits humains / Human Rights

Professeur Evelyne ASSI · Coefficient 2
DPFC / MENET-FP Côte d'Ivoire · 2025-2026

════════════════════════════════════════════════════════
  ANGLAIS LV1 · Leçon 16 — MY RIGHTS
  Niveau : 3ème | Série : Programme commun
  Professeur : Evelyne ASSI · Coefficient : 2
  Compétence : C6 · Écoute élaborée — Droits humains / Human Rights
  Programme : DPFC/MENET-FP CI 2025-2026
════════════════════════════════════════════════════════

SOMMAIRE
  Section 1 — Avant la leçon
    · Capsule de reprise (Leçon 15 : Rural exodus)
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

CAPSULE DE REPRISE — Leçon 15 : Rural exodus (L'exode rural)

5 points clés à se rappeler :
  1. L'exode rural est le déplacement massif de personnes des villages vers les villes.
  2. Les push factors sont les raisons négatives qui poussent à quitter le village (pauvreté, manque d'écoles, sécheresse).
  3. Les pull factors sont les raisons positives qui attirent vers la ville (emplois, écoles, hôpitaux).
  4. L'exode rural a des conséquences négatives pour le village (abandon, vieillissement) ET pour la ville (surpopulation, bidonvilles).
  5. Le Conditional Type 1 : IF + Present Simple + WILL + verbe base — sert à proposer des solutions.

3 questions flash :
  Q1 : What is a "slum"? (Qu'est-ce qu'un "slum" ?)
  R1 : A slum is a poor, overcrowded urban neighbourhood. (Un bidonville.)

  Q2 : Give ONE push factor of rural exodus. (Donne UN push factor de l'exode rural.)
  R2 : Poverty / lack of schools / drought. (Pauvreté / manque d'écoles / sécheresse.)

  Q3 : Complete: "If the government builds hospitals in villages, people ___ (not leave)."
  R3 : If the government builds hospitals in villages, people will not leave.

Lien avec la leçon du jour :
  Dans la leçon 15, nous avons vu que les gens quittent le village car leurs droits fondamentaux (éducation, santé) ne sont pas respectés. Aujourd'hui, nous allons apprendre à nommer et défendre ces droits en anglais — c'est le thème de la leçon 16 : My rights (Mes droits).

────────────────────────────────────────────────────────

ANCRAGE IVOIRIEN

Fatou a 14 ans. Elle vit à Abobo, une commune populaire d'Abidjan. Un jour, son instituteur lui dit : "Tu as le droit d'aller à l'école, quoi que pense ta famille." Fatou ne savait pas que l'éducation était un droit. Elle pensait que c'était un privilège réservé aux riches.

En Côte d'Ivoire, la Constitution de 2016 garantit de nombreux droits fondamentaux à tous les citoyens : le droit à l'éducation, à la santé, à la liberté d'expression, à l'égalité devant la loi. Ces droits sont également protégés par la Convention des Nations Unies relative aux droits de l'enfant (CIDE), ratifiée par la CI.

Connaître ses droits, c'est le premier pas pour les défendre.
Knowing your rights is the first step to defending them.

Lien DPFC : Ce thème est directement lié à l'éducation civique, à la citoyenneté et aux valeurs républicaines enseignées dans le programme ivoirien.

────────────────────────────────────────────────────────

▶ HARKNESS 1 — Qu'est-ce qu'un droit ?
  Q : What is a right? How is it different from a privilege?
      (Qu'est-ce qu'un droit ? En quoi est-il différent d'un privilège ?)
  Guide :
    · Un droit appartient à TOUT le monde, sans condition.
    · Un privilège est quelque chose que seules certaines personnes ont.
    · Pense à des exemples concrets : aller à l'école, voter, parler librement.
  Corrigé :
    A right is something that every person is entitled to by birth, regardless of their nationality, gender or social status. A privilege, on the other hand, is something that is given only to certain people under certain conditions. For example, going to school is a right — it belongs to every child. But being chosen to represent your school in a competition is a privilege.
    (Un droit est quelque chose auquel toute personne a droit de naissance, quelle que soit sa nationalité, son sexe ou son statut social. Un privilège, en revanche, est quelque chose qui n'est accordé qu'à certaines personnes dans certaines conditions. Par exemple, aller à l'école est un droit — il appartient à chaque enfant. Mais être choisi pour représenter son école dans une compétition est un privilège.)

▶ HARKNESS 2 — Les droits de l'enfant
  Q : What are the main rights of the child according to the United Nations?
      (Quels sont les principaux droits de l'enfant selon les Nations Unies ?)
  Guide :
    · Pense aux besoins essentiels d'un enfant : manger, apprendre, être protégé.
    · L'ONU a une convention spéciale pour les enfants — la CIDE (1989).
    · Utilise : the right to education, health, protection, identity.
  Corrigé :
    According to the United Nations Convention on the Rights of the Child (1989), every child has the right to education, healthcare, protection from violence and abuse, a name and a nationality, play and leisure, and freedom of expression. These rights apply to all children in the world, without any discrimination.
    (Selon la Convention des Nations Unies relative aux droits de l'enfant (1989), chaque enfant a droit à l'éducation, aux soins de santé, à la protection contre la violence et les mauvais traitements, à un nom et une nationalité, aux loisirs et aux jeux, et à la liberté d'expression. Ces droits s'appliquent à tous les enfants du monde, sans aucune discrimination.)

▶ HARKNESS 3 — Droits et responsabilités
  Q : Can we have rights without responsibilities?
      (Peut-on avoir des droits sans responsabilités ?)
  Guide :
    · Tout droit est accompagné d'une responsabilité correspondante.
    · Exemple : le droit à la liberté d'expression implique la responsabilité de ne pas blesser autrui.
    · Pense au lien entre droits et devoirs — thème de la prochaine leçon !
  Corrigé :
    No, rights and responsibilities always go together. If I have the right to education, I also have the responsibility to attend school regularly and respect my teachers. If I have the right to freedom of expression, I must also respect others' opinions and not spread hate speech. In a society, rights without responsibilities would lead to disorder and injustice.
    (Non, droits et responsabilités vont toujours ensemble. Si j'ai le droit à l'éducation, j'ai aussi la responsabilité d'aller régulièrement à l'école et de respecter mes professeurs. Si j'ai le droit à la liberté d'expression, je dois aussi respecter les opinions des autres et ne pas diffuser des discours haineux. Dans une société, des droits sans responsabilités conduiraient au désordre et à l'injustice.)

────────────────────────────────────────────────────────
SECTION 2 — LA LEÇON
────────────────────────────────────────────────────────

Phase 1 · Présentation / Prérequis

Titre de la leçon : My Rights (Mes droits)
Compétence : C6 — Écoute élaborée · Droits humains / Human Rights

Objectifs de la leçon :
  À la fin de cette leçon, tu seras capable de :
  · Nommer et expliquer les droits fondamentaux en anglais.
  · Distinguer les droits civils, politiques, économiques et sociaux.
  · Utiliser la Passive Voice (base) pour parler des droits garantis par la loi.
  · Exprimer l'obligation et l'interdiction avec les modaux MUST / MUST NOT / SHOULD.
  · Produire un court texte argumentatif sur les droits de l'enfant.

Prérequis nécessaires :
  · Le Present Simple (faits généraux et vérités universelles).
  · Le vocabulaire de la vie quotidienne (école, santé, famille).
  · Les modaux de base : can / cannot.
  · Notions de citoyenneté vues en Histoire-Géographie.

────────────────────────────────────────────────────────

TABLEAU DES STRUCTURES FONDAMENTALES — Leçon 16

| Expression / Notion               | Valeur, emploi                                        | Exemple anglais                                              | Traduction française                                              |
|----------------------------------|-------------------------------------------------------|--------------------------------------------------------------|-------------------------------------------------------------------|
| I have the right to + verbe      | Exprimer un droit personnel                           | I have the right to go to school.                            | J'ai le droit d'aller à l'école.                                  |
| Every child has the right to...  | Droit universel de l'enfant                           | Every child has the right to education.                      | Chaque enfant a le droit à l'éducation.                           |
| Must + verbe base                | Obligation (forte)                                    | We must respect human rights.                                | Nous devons respecter les droits humains.                         |
| Must not / mustn't               | Interdiction absolue                                  | We must not discriminate against anyone.                     | Nous ne devons pas discriminer qui que ce soit.                   |
| Should + verbe base              | Conseil / obligation morale                           | The government should protect children.                      | Le gouvernement devrait protéger les enfants.                     |
| Passive Voice (base)             | Mettre en avant l'action, pas l'auteur                | Children are protected by the law.                           | Les enfants sont protégés par la loi.                             |
| To be entitled to                | Avoir droit à (registre formel)                       | Everyone is entitled to a fair trial.                        | Tout le monde a droit à un procès équitable.                      |
| To be protected by               | Être protégé par                                      | Our rights are protected by the Constitution.                | Nos droits sont protégés par la Constitution.                     |
| To be denied                     | Se voir refuser / être nié                            | Some children are denied the right to education.             | Certains enfants se voient refuser le droit à l'éducation.        |
| Freedom of expression            | Liberté d'expression                                  | Freedom of expression is a fundamental right.                | La liberté d'expression est un droit fondamental.                 |
| Regardless of                    | Quelle que soit / sans tenir compte de                | Rights apply to everyone regardless of gender.               | Les droits s'appliquent à tous quelle que soit le genre.          |
| To violate a right               | Violer un droit                                       | Forcing a child to work violates their rights.               | Forcer un enfant à travailler viole ses droits.                   |

────────────────────────────────────────────────────────

Phase 2 · Découverte guidée

A. QU'EST-CE QU'UN DROIT ? / WHAT IS A RIGHT?

A right is something that every person is entitled to, simply because they are human.
(Un droit est quelque chose auquel toute personne a droit, simplement parce qu'elle est humaine.)

Rights are universal: they belong to everyone, everywhere, regardless of gender, race, religion or social status.
(Les droits sont universels : ils appartiennent à tout le monde, partout, quelle que soit le genre, la race, la religion ou le statut social.)

┌─ À RETENIR ──────────────────────────────────────────┐
  Rights = droits
  Human rights = droits humains / droits de l'homme
  Fundamental rights = droits fondamentaux
  Children's rights = droits de l'enfant
  Universal = universel (qui s'applique à tout le monde)
└──────────────────────────────────────────────────────┘

────────────────────────────────────────────────────────

B. LES GRANDES CATÉGORIES DE DROITS / THE MAIN CATEGORIES OF RIGHTS

1. CIVIL AND POLITICAL RIGHTS (Droits civils et politiques) :
  · The right to life (le droit à la vie)
  · The right to freedom of expression (le droit à la liberté d'expression)
  · The right to a fair trial (le droit à un procès équitable)
  · The right to vote (le droit de voter)
  · The right to freedom of religion (le droit à la liberté de religion)

2. ECONOMIC AND SOCIAL RIGHTS (Droits économiques et sociaux) :
  · The right to education (le droit à l'éducation)
  · The right to healthcare (le droit à la santé)
  · The right to work (le droit au travail)
  · The right to food and water (le droit à la nourriture et à l'eau)
  · The right to housing (le droit au logement)

3. CHILDREN'S RIGHTS (Droits de l'enfant — Convention ONU 1989) :
  · The right to a name and a nationality (le droit à un nom et une nationalité)
  · The right to education (le droit à l'éducation)
  · The right to protection from violence and abuse (le droit à la protection contre la violence)
  · The right to play and leisure (le droit au jeu et aux loisirs)
  · The right to be heard (le droit d'être entendu)

┌─ À RETENIR ──────────────────────────────────────────┐
  Ces droits sont protégés par :
  · La Déclaration Universelle des Droits de l'Homme (1948)
    = The Universal Declaration of Human Rights (UDHR)
  · La Convention des droits de l'enfant (1989)
    = The Convention on the Rights of the Child (CRC)
  · La Constitution ivoirienne de 2016
    = The Ivorian Constitution of 2016
└──────────────────────────────────────────────────────┘

────────────────────────────────────────────────────────

C. GRAMMAIRE 1 — LA PASSIVE VOICE (LA VOIX PASSIVE)

En anglais, on utilise la Passive Voice pour parler de ce qui est fait à quelqu'un ou quelque chose, sans forcément mentionner qui le fait.

C'est très courant quand on parle de droits et de lois.

FORMULE DE BASE :
  Sujet + AM / IS / ARE + participe passé (+ by + auteur si nécessaire)

Exemples :
  · Children are protected by the law.
    (Les enfants sont protégés par la loi.)
  · Human rights are guaranteed by the Constitution.
    (Les droits humains sont garantis par la Constitution.)
  · Some children are denied the right to education.
    (Certains enfants se voient refuser le droit à l'éducation.)
  · Girls are sometimes forced to get married early.
    (Les filles sont parfois forcées de se marier tôt.)

Voix active → Voix passive :
  Actif   : The government protects our rights.
            (Le gouvernement protège nos droits.)
  Passif  : Our rights are protected by the government.
            (Nos droits sont protégés par le gouvernement.)

⚠ ATTENTION
  Le participe passé des verbes irréguliers doit être mémorisé :
  · protect → protected (régulier)
  · guarantee → guaranteed (régulier)
  · give → given (irrégulier)
  · write → written (irrégulier)
  · deny → denied (régulier)
  · force → forced (régulier)

────────────────────────────────────────────────────────

D. GRAMMAIRE 2 — MUST / MUST NOT / SHOULD

Ces modaux expriment l'obligation, l'interdiction et le conseil.
Ils sont très utilisés quand on parle de droits et de devoirs.

MUST + verbe base = obligation forte (il faut absolument)
  · We must respect human rights.
    (Nous devons respecter les droits humains.)
  · Every child must go to school.
    (Chaque enfant doit aller à l'école.)

MUST NOT (MUSTN'T) + verbe base = interdiction absolue (il ne faut pas)
  · We must not discriminate against anyone.
    (Nous ne devons pas discriminer qui que ce soit.)
  · Adults must not force children to work.
    (Les adultes ne doivent pas forcer les enfants à travailler.)

SHOULD + verbe base = conseil / obligation morale (il faudrait / on devrait)
  · The government should invest in education.
    (Le gouvernement devrait investir dans l'éducation.)
  · Parents should protect their children.
    (Les parents devraient protéger leurs enfants.)

⚠ ATTENTION — Faux ami !
  "Actual" en anglais ≠ "Actuel" en français.
  · Actual (EN) = réel, véritable
  · Actuel (FR) = current (EN)
  Exemple : "The actual problem is poverty."
  = "Le véritable problème est la pauvreté." (PAS "l'actuel problème")

────────────────────────────────────────────────────────

Phase 3 · Schémas textuels

[SCHÉMA 1 — Les catégories de droits humains]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Nœud central (rectangle bleu foncé) : "MY RIGHTS / MES DROITS"
  · Branche gauche (fond bleu clair) : "CIVIL & POLITICAL RIGHTS"
      Sous-éléments : right to life · freedom of expression · right to vote · fair trial · freedom of religion
  · Branche droite (fond vert clair) : "ECONOMIC & SOCIAL RIGHTS"
      Sous-éléments : education · healthcare · work · food & water · housing
  · Branche bas (fond orange clair) : "CHILDREN'S RIGHTS (CRC 1989)"
      Sous-éléments : name & nationality · education · protection · play · be heard
  · Cadre bas (fond jaune pâle) : "PROTECTED BY"
      Sous-éléments : UDHR 1948 · CRC 1989 · Ivorian Constitution 2016
  · Flèches : de chaque branche vers le nœud central
  · Légende : bleu = droits civils | vert = droits sociaux | orange = droits enfants
Note générateur : Canva ou MindMeister — disposition en étoile autour du centre
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

[SCHÉMA 2 — Tableau comparatif : Voix active vs Voix passive]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description visuelle :
  · Deux colonnes côte à côte
  · Colonne gauche (fond vert pâle) : "ACTIVE VOICE / VOIX ACTIVE"
      Formule : Sujet + verbe + objet
      Exemple 1 : The government protects our rights.
      Exemple 2 : The law guarantees freedom.
      Exemple 3 : The teacher helps the students.
  · Colonne droite (fond bleu pâle) : "PASSIVE VOICE / VOIX PASSIVE"
      Formule : Sujet + AM/IS/ARE + participe passé (+ by + auteur)
      Exemple 1 : Our rights are protected by the government.
      Exemple 2 : Freedom is guaranteed by the law.
      Exemple 3 : The students are helped by the teacher.
  · Flèche centrale entre les deux colonnes : "TRANSFORMATION →"
  · Encadré bas : "Astuce : Passive = focus sur CE QUI EST FAIT, pas sur QUI LE FAIT"
Note générateur : tableau Word ou draw.io — 2 colonnes avec flèche centrale
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

────────────────────────────────────────────────────────

Phase 4 · Définitions DPFC

◆ DÉFINITION DPFC — Right / Un droit
  Something that every person is entitled to by birth, protected by law, and applicable to all human beings without discrimination.
  (Quelque chose auquel toute personne a droit de naissance, protégé par la loi et applicable à tous les êtres humains sans discrimination.)

◆ DÉFINITION DPFC — Human rights / Droits humains
  The basic rights and freedoms that belong to every person in the world, from birth until death, regardless of nationality, gender, religion or social status.
  (Les droits et libertés fondamentaux qui appartiennent à toute personne dans le monde, de la naissance à la mort, quelle que soit sa nationalité, son genre, sa religion ou son statut social.)

◆ DÉFINITION DPFC — Children's rights / Droits de l'enfant
  The specific rights of every child under 18 years of age, as defined by the United Nations Convention on the Rights of the Child (CRC) of 1989.
  (Les droits spécifiques de chaque enfant de moins de 18 ans, tels que définis par la Convention des Nations Unies relative aux droits de l'enfant de 1989.)

◆ DÉFINITION DPFC — Passive Voice / Voix passive
  A grammatical structure in which the subject of the sentence receives the action: Subject + AM/IS/ARE + past participle (+ by + agent).
  (Une structure grammaticale dans laquelle le sujet de la phrase reçoit l'action : Sujet + AM/IS/ARE + participe passé (+ by + auteur).)

◆ DÉFINITION DPFC — Must / Devoir (obligation)
  A modal verb expressing strong obligation or necessity. Must not expresses a strong prohibition.
  (Un verbe modal exprimant une obligation forte ou une nécessité. Must not exprime une interdiction absolue.)

◆ DÉFINITION DPFC — Discrimination
  The unfair treatment of a person or group based on characteristics such as gender, race, religion or social status.
  (Le traitement injuste d'une personne ou d'un groupe fondé sur des caractéristiques telles que le genre, la race, la religion ou le statut social.)
  (En anglais : Discrimination — même orthographe, même sens.)

────────────────────────────────────────────────────────

Phase 5 · Trace écrite + Analogies CI

✎ TRACE ÉCRITE (à recopier dans le cahier)

I. VOCABULAIRE CLÉS — MY RIGHTS / MES DROITS

  Right = droit
  Human rights = droits humains
  Children's rights = droits de l'enfant
  Fundamental = fondamental
  Universal = universel
  The right to education = le droit à l'éducation
  The right to healthcare = le droit à la santé
  The right to life = le droit à la vie
  Freedom of expression = liberté d'expression
  The right to vote = le droit de voter
  The right to a fair trial = le droit à un procès équitable
  The right to housing = le droit au logement
  To be entitled to = avoir droit à
  To be protected by = être protégé par
  To be denied = se voir refuser
  To violate = violer (un droit)
  Discrimination = discrimination
  Regardless of = quelle que soit / sans tenir compte de
  The UDHR (1948) = La Déclaration Universelle des Droits de l'Homme
  The CRC (1989) = La Convention des droits de l'enfant

II. CATÉGORIES DE DROITS / CATEGORIES OF RIGHTS

  1. Civil and political rights (droits civils et politiques) :
     right to life · freedom of expression · right to vote · fair trial

  2. Economic and social rights (droits économiques et sociaux) :
     education · healthcare · work · food · housing

  3. Children's rights (droits de l'enfant) :
     name & nationality · education · protection · play · be heard

III. GRAMMAIRE — PASSIVE VOICE (VOIX PASSIVE)

  Formule : Sujet + AM / IS / ARE + participe passé (+ by + auteur)
  Exemple 1 : Children are protected by the law.
              (Les enfants sont protégés par la loi.)
  Exemple 2 : Human rights are guaranteed by the Constitution.
              (Les droits humains sont garantis par la Constitution.)
  Exemple 3 : Some girls are denied the right to education.
              (Certaines filles se voient refuser le droit à l'éducation.)

IV. GRAMMAIRE — MUST / MUST NOT / SHOULD

  MUST + verbe base = obligation forte
    We must respect human rights.
    (Nous devons respecter les droits humains.)

  MUST NOT (MUSTN'T) + verbe base = interdiction absolue
    We must not discriminate against anyone.
    (Nous ne devons pas discriminer qui que ce soit.)

  SHOULD + verbe base = conseil / obligation morale
    The government should protect children.
    (Le gouvernement devrait protéger les enfants.)

V. FAUX AMI À RETENIR

  Actual (EN) ≠ Actuel (FR)
  Actual = réel, véritable
  Current = actuel
  Exemple : "The actual problem is poverty." = "Le véritable problème est la pauvreté."

────────────────────────────────────────────────────────

~ ANALOGIE CI 1
  Les droits humains, c'est comme les règles du jeu d'awalé : elles existent pour que tout le monde joue de façon juste. Si quelqu'un ne respecte pas les règles, le jeu devient injuste. De même, si les droits ne sont pas respectés, la société devient injuste.

~ ANALOGIE CI 2
  La Passive Voice, c'est comme quand on dit au marché d'Adjamé : "Les pagnes sont vendus ici" au lieu de "Le commerçant vend les pagnes ici." On met l'accent sur ce qui est fait (les pagnes sont vendus), pas sur qui le fait.

────────────────────────────────────────────────────────

Phase 5b · Synthèse

5 points essentiels à retenir :
  1. Un droit est universel : il appartient à tout être humain, sans condition ni discrimination.
  2. Il existe trois grandes catégories : droits civils et politiques · droits économiques et sociaux · droits de l'enfant.
  3. Ces droits sont protégés par la DUDH (1948), la CRC (1989) et la Constitution ivoirienne (2016).
  4. La Passive Voice (AM/IS/ARE + participe passé) est utilisée pour parler de ce qui est garanti ou violé par la loi.
  5. MUST exprime l'obligation forte · MUST NOT l'interdiction · SHOULD le conseil moral.

5 mots clés + définition courte :
  · Right : quelque chose auquel tout être humain a droit de naissance.
  · Universal : qui s'applique à tout le monde, sans exception.
  · Passive Voice : structure grammaticale où le sujet reçoit l'action (IS/ARE + participe passé).
  · Must not : modal exprimant une interdiction absolue.
  · Discrimination : traitement injuste basé sur la race, le genre, la religion, etc.

3 erreurs fréquentes + correction :
  Erreur 1 : Écrire "Rights is important." (oubli du pluriel avec ARE)
  ✔ Correction : "Rights are important." — rights est pluriel → ARE.

  Erreur 2 : Écrire "The rights are protect by the law." (oubli du participe passé)
  ✔ Correction : "The rights are protected by the law." — il faut le participe passé : protected.

  Erreur 3 : Confondre MUST NOT et SHOULD NOT.
  ✔ Correction : MUST NOT = interdiction absolue (il ne faut JAMAIS). SHOULD NOT = conseil négatif (il vaudrait mieux ne pas). Exemple : "You must not steal." ≠ "You should not eat too much sugar."

────────────────────────────────────────────────────────

Phase 6 · Exercices guidés

◉ EXERCICE GUIDÉ 1 — Active to Passive · Notion ciblée : Passive Voice

  Énoncé :
  Transform the following sentences from active voice to passive voice.
  Transforme les phrases suivantes de la voix active à la voix passive.

  1. The law protects every child.
  2. The government guarantees freedom of expression.
  3. Teachers educate young people.

  MÉTHODE
  Étape 1 · Identifie le sujet, le verbe et l'objet de la phrase active.
            Sujet → Verbe → Objet
  Étape 2 · L'objet de la phrase active devient le SUJET de la phrase passive.
  Étape 3 · Conjugue le verbe ÊTRE (AM/IS/ARE) en accord avec le nouveau sujet.
  Étape 4 · Écris le participe passé du verbe principal.
  Étape 5 · Ajoute "by + ancien sujet" si nécessaire.
  ✔ Vérification : Vérifie que AM/IS/ARE s'accorde avec le nouveau sujet (singulier/pluriel).
  ✔ Conclusion :
    1. Every child is protected by the law.
       (Chaque enfant est protégé par la loi.)
    2. Freedom of expression is guaranteed by the government.
       (La liberté d'expression est garantie par le gouvernement.)
    3. Young people are educated by teachers.
       (Les jeunes sont éduqués par les enseignants.)
  ✔ Ce que cet exercice mobilise : Passive Voice au Present Simple · accord de AM/IS/ARE · participes passés · transformation de structure de phrase · traduction anglais-français.

────────────────────────────────────────────────────────

◉ EXERCICE GUIDÉ 2 — Must / Must not / Should · Notion ciblée : Modaux d'obligation et de conseil

  Énoncé :
  Fill in the blanks with MUST, MUST NOT or SHOULD.
  Complète les blancs avec MUST, MUST NOT ou SHOULD.

  1. Every child ___ go to school. It is the law.
  2. Adults ___ force children to work. It is forbidden.
  3. The government ___ invest more in rural schools. It would really help.
  4. We ___ discriminate against girls. It is a violation of their rights.
  5. Parents ___ listen to their children's needs and opinions.

  MÉTHODE
  Étape 1 · Lis chaque phrase et identifie le type de message : obligation forte, interdiction ou conseil ?
  Étape 2 · Obligation forte (il faut absolument) → MUST.
  Étape 3 · Interdiction absolue (il ne faut pas) → MUST NOT.
  Étape 4 · Conseil / obligation morale (il faudrait) → SHOULD.
  Étape 5 · Vérifie que le verbe après le modal est toujours à la forme de base (sans -s, sans -ed).
  ✔ Vérification : Relis chaque phrase complétée pour vérifier le sens.
  ✔ Conclusion :
    1. MUST — c'est une obligation légale.
       Every child must go to school.
       (Chaque enfant doit aller à l'école.)
    2. MUST NOT — c'est une interdiction absolue.
       Adults must not force children to work.
       (Les adultes ne doivent pas forcer les enfants à travailler.)
    3. SHOULD — c'est un conseil, une recommandation.
       The government should invest more in rural schools.
       (Le gouvernement devrait investir davantage dans les écoles rurales.)
    4. MUST NOT — c'est une interdiction absolue.
       We must not discriminate against girls.
       (Nous ne devons pas discriminer les filles.)
    5. SHOULD — c'est une obligation morale / un conseil.
       Parents should listen to their children's needs and opinions.
       (Les parents devraient écouter les besoins et les opinions de leurs enfants.)
  ✔ Ce que cet exercice mobilise : modaux MUST / MUST NOT / SHOULD · distinction obligation / interdiction / conseil · verbe base après modal · traduction anglais-français · vocabulaire des droits.

────────────────────────────────────────────────────────
SECTION 3 — APRÈS LA LEÇON
────────────────────────────────────────────────────────

◎ EXERCICE 1 — Vocabulary matching · Notions travaillées : Vocabulaire des droits humains

  Match each English expression with its French translation.
  Relie chaque expression anglaise à sa traduction française.

  1. Freedom of expression         a. Droit au logement
  2. The right to a fair trial     b. Se voir refuser
  3. To be entitled to             c. Liberté d'expression
  4. The right to housing          d. Avoir droit à
  5. To be denied                  e. Droit à un procès équitable

  GUIDE
  Étape 1 · Relis la trace écrite, section I (Vocabulaire clés).
  Étape 2 · Pour chaque expression anglaise, cherche sa traduction dans ta liste.
  Étape 3 · Relie par une flèche ou écris la lettre correspondante.

────────────────────────────────────────────────────────

◎ EXERCICE 2 — Classifying rights · Notions travaillées : Catégories de droits

  Classify the following rights into the correct category.
  Classe les droits suivants dans la bonne catégorie.

  Rights to classify :
    the right to vote · the right to education · the right to play · the right to food ·
    the right to a fair trial · the right to a name · freedom of religion · the right to healthcare

  Categories :
    A. Civil and political rights (Droits civils et politiques)
    B. Economic and social rights (Droits économiques et sociaux)
    C. Children's rights (Droits de l'enfant)

  GUIDE
  Étape 1 · Relis la trace écrite, section II (Catégories de droits).
  Étape 2 · Pour chaque droit, demande-toi : est-ce un droit lié à la vie politique ? à la vie sociale ? ou spécifique aux enfants ?
  Étape 3 · Note la lettre (A, B ou C) en face de chaque droit.
  Étape 4 · Certains droits peuvent appartenir à deux catégories — justifie ton choix.

────────────────────────────────────────────────────────

◎ EXERCICE 3 — Passive Voice · Notions travaillées : Transformation active → passive

  Rewrite the following sentences in the passive voice.
  Réécris les phrases suivantes à la voix passive.

  1. The Constitution guarantees our fundamental rights.
  2. The United Nations protects children all over the world.
  3. Society denies some girls the right to education.
  4. The law forbids discrimination based on gender.
  5. The government must provide healthcare to all citizens.

  GUIDE
  Étape 1 · Identifie sujet + verbe + objet dans chaque phrase active.
  Étape 2 · L'objet devient le nouveau sujet de la phrase passive.
  Étape 3 · Conjugue IS ou ARE selon le nouveau sujet (singulier ou pluriel).
  Étape 4 · Écris le participe passé du verbe principal.
  Étape 5 · Ajoute "by + ancien sujet" si nécessaire.
  Étape 6 · Pour la phrase 5 : utilise MUST BE + participe passé (modal + passive).
  Étape 7 · Traduis chaque phrase en français.

────────────────────────────────────────────────────────

◎ EXERCICE 4 — Must / Must not / Should · Notions travaillées : Modaux d'obligation, d'interdiction et de conseil

  Read each situation and write a sentence using MUST, MUST NOT or SHOULD.
  Lis chaque situation et écris une phrase avec MUST, MUST NOT ou SHOULD.

  1. A child is forced to work in a field instead of going to school.
     (Start with: "Children...")
  2. A girl is refused entry to school because she is a girl.
     (Start with: "Schools...")
  3. A government does not build enough hospitals in rural areas.
     (Start with: "The government...")
  4. A man is put in prison without a trial.
     (Start with: "No one...")
  5. People in a village have no access to clean water.
     (Start with: "Everyone...")

  GUIDE
  Étape 1 · Identifie si la situation décrit quelque chose d'interdit, d'obligatoire ou qu'on devrait faire.
  Étape 2 · Choisis le bon modal : MUST (obligation), MUST NOT (interdiction), SHOULD (conseil).
  Étape 3 · Écris ta phrase avec le modal + verbe base.
  Étape 4 · Traduis chaque phrase en français.

────────────────────────────────────────────────────────

◎ EXERCICE 5 — Short comprehension · Notions travaillées : Compréhension écrite + vocabulaire des droits

  Read the text and answer the questions in English.
  Lis le texte et réponds aux questions en anglais.

  TEXT:
  "Mariama is 13 years old. She lives in a village in the west of Côte d'Ivoire. Her parents want her to get married before she finishes school. Mariama says: 'I have the right to education. I have the right to choose my own life. No one should force me to marry before I am ready. The law must protect me.' Her teacher agrees and contacts a local organization that helps children know and defend their rights."

  Questions :
  1. How old is Mariama?
  2. What right is Mariama defending?
  3. Find ONE sentence in the passive voice in the text. (Attention : look carefully.)
  4. Which modal does Mariama use to express a strong obligation? Give the sentence.
  5. Do you think Mariama is right? Write TWO sentences to give your opinion in English.

  GUIDE
  Étape 1 · Lis le texte une première fois pour comprendre la situation générale.
  Étape 2 · Pour la question 3 : cherche une phrase avec IS/ARE + participe passé.
  Étape 3 · Pour la question 4 : rappelle-toi que MUST = obligation forte.
  Étape 4 · Pour la question 5 : utilise "I think that..." ou "In my opinion..." pour donner ton avis.
  Étape 5 · Réponds en anglais avec des phrases complètes.

────────────────────────────────────────────────────────

◈ DEVOIR MAISON — Defending my rights
  Durée conseillée : 35 minutes

  Sujet :
  Write a short text of 8 to 10 sentences in English.
  (Écris un court texte de 8 à 10 phrases en anglais.)

  You are a student at a school in Abidjan. You have noticed that:
  · Some girls in your class are sometimes prevented from speaking in class.
  · Some students do not have access to textbooks.
  · A student was punished unfairly without explanation.

  Write a letter to your school principal (directeur) to defend the rights of your classmates.
  Écris une lettre à ton directeur d'école pour défendre les droits de tes camarades.

  Your letter must include :
  · A greeting (Dear Mr/Ms...)
  · The rights that are being violated (use Passive Voice : "...are denied / are not respected...")
  · What the school MUST and MUST NOT do (use MUST / MUST NOT)
  · What you think SHOULD change (use SHOULD)
  · A polite closing (Yours sincerely, / Respectfully yours,)

  Contraintes obligatoires :
  · 8 à 10 phrases minimum.
  · Au moins 2 phrases à la Passive Voice.
  · Au moins une phrase avec MUST, une avec MUST NOT et une avec SHOULD.
  · Utiliser au moins 5 mots du vocabulaire de la leçon.
  · Traduire la lettre en français après l'avoir écrite en anglais.

  GUIDE (sans corrigé)
  Étape 1 · Relis la trace écrite : vocabulaire des droits, Passive Voice, MUST/MUST NOT/SHOULD.
  Étape 2 · Au brouillon, note les 3 problèmes et les droits correspondants violés.
  Étape 3 · Rédige ta lettre : salutation → problèmes (passive) → obligations (must) → conseils (should) → conclusion.
  Étape 4 · Vérifie : passive = IS/ARE + participe passé. Modal = verbe base sans -s.
  Étape 5 · Traduis ta lettre en français.

────────────────────────────────────────────────────────
SECTION 4 — CORRIGÉS COMPLETS
────────────────────────────────────────────────────────

✔ CORRIGÉ — DEVOIR MAISON : Defending my rights

Proposition de corrigé (lettre modèle en anglais) :

Dear Mr Principal,

I am a student in class 3ème B. I am writing to inform you about some problems that affect the rights of my classmates.

First, some girls in our class are sometimes prevented from speaking during lessons. This is not fair. The right to be heard is guaranteed to every student, regardless of gender.

Second, many students are not provided with textbooks. We cannot study properly without books. Every student must be given access to learning materials.

Third, one of our classmates was punished last week without any explanation. No student must be punished without a fair reason. The right to a fair process must be respected, even at school.

I think our school should review its rules to make sure all students are treated equally. Girls and boys should have the same opportunities to speak and learn.

I hope you will take action to improve the situation.

Yours sincerely,
Kouassi Ange, 3ème B

Traduction française complète :
Cher Monsieur le Directeur,

Je suis élève en classe de 3ème B. Je vous écris pour vous informer de certains problèmes qui affectent les droits de mes camarades.

Premièrement, certaines filles de notre classe sont parfois empêchées de s'exprimer pendant les cours. Ce n'est pas juste. Le droit d'être entendu est garanti à chaque élève, quel que soit le genre.

Deuxièmement, de nombreux élèves ne reçoivent pas de manuels scolaires. Nous ne pouvons pas étudier correctement sans livres. Chaque élève doit avoir accès aux supports d'apprentissage.

Troisièmement, l'un de nos camarades a été puni la semaine dernière sans aucune explication. Aucun élève ne doit être puni sans raison valable. Le droit à un traitement équitable doit être respecté, même à l'école.

Je pense que notre école devrait revoir ses règles pour s'assurer que tous les élèves sont traités de façon égale. Les filles et les garçons devraient avoir les mêmes possibilités de s'exprimer et d'apprendre.

J'espère que vous prendrez des mesures pour améliorer la situation.

Respectueusement,
Kouassi Ange, 3ème B

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 1 : Vocabulary matching

  1. Freedom of expression → c. Liberté d'expression
  2. The right to a fair trial → e. Droit à un procès équitable
  3. To be entitled to → d. Avoir droit à
  4. The right to housing → a. Droit au logement
  5. To be denied → b. Se voir refuser

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 2 : Classifying rights

  A. Civil and political rights (Droits civils et politiques) :
     · The right to vote (le droit de voter)
     · The right to a fair trial (le droit à un procès équitable)
     · Freedom of religion (la liberté de religion)

  B. Economic and social rights (Droits économiques et sociaux) :
     · The right to education (le droit à l'éducation)
     · The right to food (le droit à la nourriture)
     · The right to healthcare (le droit à la santé)

  C. Children's rights (Droits de l'enfant) :
     · The right to play (le droit au jeu)
     · The right to a name (le droit à un nom)

  Note : The right to education peut aussi être classé en C (droits de l'enfant) — les deux réponses sont acceptées si justifiées.

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 3 : Passive Voice

  1. Our fundamental rights are guaranteed by the Constitution.
     (Nos droits fondamentaux sont garantis par la Constitution.)

  2. Children all over the world are protected by the United Nations.
     (Les enfants du monde entier sont protégés par les Nations Unies.)

  3. Some girls are denied the right to education by society.
     (Certaines filles se voient refuser le droit à l'éducation par la société.)

  4. Discrimination based on gender is forbidden by the law.
     (La discrimination basée sur le genre est interdite par la loi.)

  5. Healthcare must be provided to all citizens by the government.
     (Les soins de santé doivent être fournis à tous les citoyens par le gouvernement.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 4 : Must / Must not / Should

  1. Children must not be forced to work. They must go to school.
     (Les enfants ne doivent pas être forcés à travailler. Ils doivent aller à l'école.)

  2. Schools must not refuse entry to girls. Education is a right for everyone.
     (Les écoles ne doivent pas refuser l'entrée aux filles. L'éducation est un droit pour tous.)

  3. The government should build more hospitals in rural areas.
     (Le gouvernement devrait construire plus d'hôpitaux dans les zones rurales.)

  4. No one must be put in prison without a fair trial.
     (Personne ne doit être mis en prison sans procès équitable.)

  5. Everyone should have access to clean water. It is a basic right.
     (Tout le monde devrait avoir accès à l'eau potable. C'est un droit fondamental.)

────────────────────────────────────────────────────────

✔ CORRIGÉ — EXERCICE 5 : Short comprehension

  1. Mariama is 13 years old.
     (Mariama a 13 ans.)

  2. Mariama is defending her right to education and her right to choose her own life / the right not to be forced into marriage.
     (Mariama défend son droit à l'éducation et son droit de choisir sa propre vie / le droit de ne pas être forcée à se marier.)

  3. Passive Voice in the text : "The law must protect me." — Attention : cette phrase n'est pas à la passive. La phrase passive est implicitement : "No one should force me" contient une structure négative avec modal. La vraie réponse attendue : il n'y a pas de passive Voice explicite avec IS/ARE dans le texte — l'élève qui identifie cela correctement mérite les points. Réponse possible acceptable : "No one should force me to marry" → peut être reformulé en passif : "I should not be forced to marry."
     Note pédagogique : cet exercice développe l'esprit critique — le texte ne contient pas de passive Voice classique, ce qui pousse l'élève à bien analyser les structures.

  4. Mariama uses MUST to express a strong obligation : "The law must protect me."
     (Mariama utilise MUST pour exprimer une obligation forte : "La loi doit me protéger.")

  5. Proposition de réponse :
     Yes, I think Mariama is right. Every child has the right to education and to choose their own future. No one should be forced to marry at the age of 13.
     (Oui, je pense que Mariama a raison. Chaque enfant a le droit à l'éducation et de choisir son propre avenir. Personne ne devrait être forcé à se marier à l'âge de 13 ans.)

════════════════════════════════════════════════════════
  CITATION FINALE
════════════════════════════════════════════════════════

  "Where, after all, do universal human rights begin? In small places, close to home."
  (Où commencent, après tout, les droits humains universels ? Dans de petits endroits, proches de chez soi.)

  — Eleanor Roosevelt, 1958

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
