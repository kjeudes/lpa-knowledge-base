# STYLE_MASTER_LYCEE.md — Styles LPA v9.1 · CI 2025-2026

## COULEURS PAR NIVEAU

```
3ème      → Vert CI   #2E7D32
Seconde   → Ardoise   #1565C0
Première  → Bordeaux  #6A1B1A
Terminale → Noir      #1A1A1A + Or #C4972F
```

## COULEURS PAR MATIÈRE

```
Français        → Bordeaux    #8B1A1A  · light #F9ECEC
Mathématiques   → Bleu        #1565C0  · light #E8F0FB
Histoire-Géo    → Ocre        #8B6914  · light #F7F0E0
Anglais LV1     → Vert foncé  #1B5E20  · light #E8F5E9
SVT             → Vert        #2E7D32  · light #EAF4EA
Physique-Chimie → Violet      #4A148C  · light #F0EAF8
Espagnol LV2    → Orange      #E65100  · light #FEF0E8
Philosophie     → Ardoise     #37474F  · light #EEF0F1
Maths + PC schémas : a>0 Bleu #1565C0 · a<0 Orange #E65100
```

---

## PRINCIPE v9.1 — CSS INTÉGRÉ DEPUIS STYLE_MASTER

```
Claude copie INTÉGRALEMENT le bloc CSS_LPA ci-dessous dans <style> de chaque leçon.
Claude ajuste UNIQUEMENT data-matiere sur <body> selon la matière.
Aucune autre modification du CSS n'est autorisée.
```

---

## CORRESPONDANCE BALISES → STYLES

### Encadrés (aside[data-type])
```
aside[data-type="retenir"]    → border-left matière · fond light matière
aside[data-type="attention"]  → border-left orange · fond #FFF3E0
aside[data-type="definition"] → border matière · fond blanc · titre centré
aside[data-type="trace"]      → border dashed matière · fond #FAFAFA
aside[data-type="analogie"]   → border-left vert · fond #F1F8E9
aside[data-type="pont"]       → border-left violet · fond #F3E5F5
aside[data-type="numerique"]  → border dashed bleu · fond #E3F2FD
```

### Autres éléments
```
div.harkness → p.hq (fond matière) · p.hg (fond jaune) · p.hc (fond vert)
ol.guide > li → compteur CSS cercle matière · em = raison italique gris
ol.corrige-etapes > li → compteur CSS cercle gris
article.phase · article.exercice-guide · article.exercice · article.devoir · article.corrige
p.verif · p.conclu · figure · figcaption · table · th · td
```

---

## RÈGLES FORMAT ABSOLUES

```
→ Numérotation : <ol> + compteur CSS — jamais de numéros manuels
→ Tableaux : largeur 100% · colonnes équilibrées
→ Schémas : SVG inline dans <figure> — jamais PNG externe ni ASCII
→ Aucun attribut style="" — jamais
→ Classes autorisées uniquement : page · page-garde · phase · harkness · hq · hg · hc
  guide · verif · conclu · exercice-guide · exercice · devoir · corrige · corrige-etapes
```

---

## STYLES PAR NIVEAU (ton pédagogique)

```
3ème      → ton encourageant · vocabulaire accessible · "Au BEPC, on te demandera..."
Seconde   → ton structuré · notions nouvelles expliquées · "En vue du BAC, retenez que..."
Première  → ton exigeant · vocabulaire académique · "Le BAC exige que vous sachiez..."
Terminale → ton expert · sujets BAC réels · "Sujet type BAC CI..."
```

---

## CSS_LPA — BLOC À COPIER INTÉGRALEMENT DANS CHAQUE LEÇON

```css
:root{--vert:#2E7D32;--orange:#E65100;--r:8px;--s:0 2px 8px rgba(0,0,0,.08)}
[data-matiere="maths"]{--c:#1565C0;--cl:#E8F0FB}
[data-matiere="francais"]{--c:#8B1A1A;--cl:#F9ECEC}
[data-matiere="hg"]{--c:#8B6914;--cl:#F7F0E0}
[data-matiere="anglais"]{--c:#1B5E20;--cl:#E8F5E9}
[data-matiere="svt"]{--c:#2E7D32;--cl:#EAF4EA}
[data-matiere="pc"]{--c:#4A148C;--cl:#F0EAF8}
[data-matiere="espagnol"]{--c:#E65100;--cl:#FEF0E8}
[data-matiere="philo"]{--c:#37474F;--cl:#EEF0F1}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:Arial,sans-serif;font-size:11pt;line-height:1.6;color:#1A1A1A;background:#F5F5F5}
div.page{max-width:780px;margin:0 auto;background:white;padding:2cm}
header.page-garde{min-height:100vh;display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center;border-top:8px solid var(--c);page-break-after:always}
header.page-garde .lpa-logo{font-size:2.5em;font-weight:bold;color:var(--c);letter-spacing:4px;margin-bottom:2rem}
header.page-garde h1{font-size:2em;color:var(--c);margin-bottom:.5rem}
header.page-garde p{font-size:1em;color:#555;margin-bottom:.3rem}
header.page-garde .dpfc{margin-top:2rem;font-size:.9em;color:#888}
nav#sommaire{background:var(--cl);border-left:5px solid var(--c);border-radius:var(--r);padding:1.5rem;margin:2rem 0;page-break-after:always}
nav#sommaire h2{color:var(--c);font-size:1.2em;margin-bottom:1rem;text-transform:uppercase;letter-spacing:2px;background:none;padding:0}
nav#sommaire ol,nav#sommaire ul{columns:2;column-gap:2rem;list-style:none;padding:0}
nav#sommaire li{break-inside:avoid;margin-bottom:.3rem}
nav#sommaire a{color:#333;text-decoration:none;font-size:.9em;display:block;padding:.1rem 0}
nav#sommaire a:hover{color:var(--c)}
nav#sommaire strong{color:var(--c);display:block;margin:.8rem 0 .3rem}
section>h2{background:var(--c);color:white;padding:1rem 1.5rem;border-radius:var(--r) var(--r) 0 0;margin:2.5rem 0 0;font-size:1.3em}
article.phase{border:1px solid #E0E0E0;border-radius:0 0 var(--r) var(--r);padding:1.5rem;margin-bottom:1.5rem;box-shadow:var(--s)}
article.phase>h3{font-size:1.15em;color:#1A1A1A;border-bottom:2px solid var(--cl);padding-bottom:.5rem;margin-bottom:1rem}
article.phase h4,h4{color:var(--c);font-size:1em;margin:1rem 0 .5rem}
h3{color:var(--c);margin:1.2rem 0 .5rem}
p{margin:.6rem 0}
ul,ol{padding-left:1.5rem;margin:.5rem 0}
strong{color:var(--c)}
em{font-style:italic}
code{background:#F5F5F5;border:1px solid #DDD;padding:.1rem .4rem;border-radius:3px;font-family:monospace;font-size:.95em}
aside{padding:1rem 1.2rem;border-radius:0 var(--r) var(--r) 0;margin:1rem 0}
aside>strong:first-child{display:block;margin-bottom:.4rem;font-size:.85em;text-transform:uppercase;letter-spacing:1px}
aside[data-type="retenir"]{border-left:4px solid var(--c);background:var(--cl)}
aside[data-type="retenir"]>strong:first-child{color:var(--c)}
aside[data-type="attention"]{border-left:4px solid var(--orange);background:#FFF3E0}
aside[data-type="attention"]>strong:first-child{color:var(--orange)}
aside[data-type="definition"]{border:1px solid var(--c);background:white;border-radius:var(--r)}
aside[data-type="definition"]>strong:first-child{text-align:center;color:var(--c);display:block}
aside[data-type="trace"]{border:2px dashed var(--c);background:#FAFAFA;border-radius:var(--r)}
aside[data-type="trace"]>strong:first-child{color:var(--c)}
aside[data-type="analogie"]{border-left:4px solid var(--vert);background:#F1F8E9}
aside[data-type="analogie"]>strong:first-child{color:var(--vert)}
aside[data-type="pont"]{border-left:4px solid #6A1B9A;background:#F3E5F5}
aside[data-type="pont"]>strong:first-child{color:#6A1B9A}
aside[data-type="numerique"]{border:1px dashed #90CAF9;background:#E3F2FD;border-radius:var(--r);color:#1565C0}
div.harkness{border:1px solid #E0E0E0;border-radius:var(--r);margin:1rem 0;overflow:hidden;box-shadow:var(--s)}
p.hq{background:var(--c);color:white;padding:.8rem 1rem;font-weight:bold;margin:0}
p.hg{background:#FFFDE7;padding:.8rem 1rem;border-bottom:1px solid #E0E0E0;font-size:.95em;margin:0}
p.hc{background:#F1F8E9;padding:.8rem 1rem;font-size:.95em;margin:0}
p.hc strong{color:var(--vert)}
ol.guide{list-style:none;padding:0;counter-reset:guide-counter}
ol.guide>li{counter-increment:guide-counter;display:flex;gap:1rem;padding:.6rem 0;border-bottom:1px solid #F0F0F0;align-items:flex-start;flex-wrap:wrap}
ol.guide>li::before{content:counter(guide-counter);background:var(--c);color:white;font-weight:bold;font-size:.85em;padding:.2rem .6rem;border-radius:50%;min-width:28px;text-align:center;flex-shrink:0;line-height:1.6}
ol.guide>li>em{display:block;font-style:italic;color:#555;font-size:.9em;margin-top:.2rem;flex-basis:100%;padding-left:44px}
p.verif{background:#E8F5E9;border-left:3px solid var(--vert);padding:.5rem .8rem;margin-top:.8rem;font-size:.9em;border-radius:0 4px 4px 0}
p.conclu{background:var(--cl);border-radius:var(--r);padding:.5rem .8rem;margin-top:.8rem;font-weight:bold;font-size:.9em}
article.exercice-guide{border-left:4px solid var(--vert);background:#F1F8E9;border-radius:0 var(--r) var(--r) 0;padding:1.2rem;margin:1rem 0}
article.exercice-guide h4{color:var(--vert);margin-bottom:.8rem}
article.exercice{border:1px solid #E0E0E0;border-radius:var(--r);padding:1.2rem;margin:1rem 0;box-shadow:0 1px 4px rgba(0,0,0,.06)}
article.exercice h4{color:var(--c)}
article.devoir{border-left:4px solid var(--c);background:var(--cl);border-radius:0 var(--r) var(--r) 0;padding:1.2rem;margin:1.5rem 0}
article.devoir h4{color:var(--c)}
article.corrige{background:#F9F9F9;border:1px solid #E0E0E0;border-radius:var(--r);padding:1.5rem;margin-top:1.5rem}
article.corrige h4{color:var(--c);margin-bottom:.8rem}
ol.corrige-etapes{list-style:none;padding:0;counter-reset:corrige-counter}
ol.corrige-etapes>li{counter-increment:corrige-counter;display:flex;gap:1rem;padding:.5rem 0;align-items:flex-start}
ol.corrige-etapes>li::before{content:counter(corrige-counter);background:#555;color:white;font-size:.8em;padding:.15rem .5rem;border-radius:50%;min-width:26px;text-align:center;flex-shrink:0;line-height:1.6}
figure{margin:1rem 0;text-align:center}
figure svg{width:100%}
figcaption{font-size:.85em;color:#555;margin-top:.5rem;font-style:italic}
table{width:100%;border-collapse:collapse;margin:1rem 0}
th{background:var(--c);color:white;padding:.6rem .8rem;text-align:left;font-size:.9em}
td{padding:.5rem .8rem;border-bottom:1px solid #E0E0E0;font-size:.9em}
tr:nth-child(even) td{background:var(--cl)}
@media print{body{background:white}div.page{max-width:100%;padding:1cm;box-shadow:none}header.page-garde{page-break-after:always}nav#sommaire{page-break-after:always}article.phase{box-shadow:none}a{color:inherit;text-decoration:none}}
```
