require('dotenv').config();
const express = require('express');
const cors    = require('cors');
const path    = require('path');
const {
  Document, Packer, Paragraph, TextRun,
  HeadingLevel, AlignmentType, BorderStyle,
  ShadingType, TableRow, TableCell, Table, WidthType
} = require('docx');

const app     = express();
const PORT    = process.env.PORT || 3000;
const API_KEY = process.env.ANTHROPIC_API_KEY;
const ANTHROPIC_URL = 'https://api.anthropic.com/v1/messages';

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── Référentiels professionnels ───────────────────────────────────────────
const REFERENTIELS = {
  'bac-pro-mdcd': {
    label: 'Bac Pro Métiers du Commerce et de la Distribution (MDCD)',
    classes: ['2nde Bac Pro', '1ère Bac Pro', 'Terminale Bac Pro'],
    activites: {
      'option-a': {
        label: "Option A — Gestion des stocks et de l'espace commercial",
        competences: [
          'Réceptionner les livraisons',
          'Contrôler les livraisons',
          'Contribuer à la gestion des stocks',
          'Participer aux inventaires',
          'Mettre en œuvre des actions promotionnelles',
          "Gérer l'espace commercial au quotidien",
        ]
      },
      'option-b': {
        label: 'Option B — Prospection et vente à distance',
        competences: [
          'Préparer et organiser la prospection',
          'Prospecter une clientèle de particuliers ou de professionnels',
          'Analyser les résultats de la prospection',
          "Valoriser l'offre commerciale",
          "Négocier et défendre l'offre",
        ]
      }
    }
  },
  'cap-epc': {
    label: 'CAP Équipier Polyvalent du Commerce (EPC)',
    classes: ['CAP 1ère année', 'CAP 2ème année'],
    activites: {
      'reception': {
        label: 'Réception et contrôle des marchandises',
        competences: [
          'Réceptionner les livraisons',
          'Contrôler la conformité des marchandises',
          'Étiqueter et conditionner les produits',
          'Ranger et stocker les marchandises',
        ]
      },
      'rayon': {
        label: "Mise en rayon et tenue de l'espace commercial",
        competences: [
          'Approvisionner les rayons',
          'Implanter et baliser les produits',
          "Maintenir la propreté et l'ordre du rayon",
          'Gérer les dates limites de consommation',
        ]
      },
      'vente': {
        label: 'Vente et relation client',
        competences: [
          'Accueillir et orienter les clients',
          'Informer et conseiller les clients',
          'Réaliser les opérations de caisse',
          'Traiter les réclamations courantes',
        ]
      }
    }
  }
};

// ── Mapping compétence → document professionnel recommandé ───────────────
const COMPETENCE_DOCUMENTS = {
  // Bac Pro MDCD — Option A
  'Réceptionner les livraisons':                                    'BON DE LIVRAISON',
  'Contrôler les livraisons':                                       'BON DE RÉCEPTION',
  'Contribuer à la gestion des stocks':                             'FICHE DE STOCK',
  'Participer aux inventaires':                                     "FICHE D'INVENTAIRE",
  'Mettre en œuvre des actions promotionnelles':                    'BON DE COMMANDE',
  "Gérer l'espace commercial au quotidien":                         'RELEVÉ DE DÉMARQUE',
  // Bac Pro MDCD — Option B
  'Préparer et organiser la prospection':                           'FICHE DÉCOUVERTE CLIENT',
  'Prospecter une clientèle de particuliers ou de professionnels':  'FICHE DÉCOUVERTE CLIENT',
  'Analyser les résultats de la prospection':                       'FICHE DÉCOUVERTE CLIENT',
  "Valoriser l'offre commerciale":                                  'BON DE COMMANDE CLIENT',
  "Négocier et défendre l'offre":                                   'BON DE COMMANDE CLIENT',
  // CAP EPC — Réception
  'Contrôler la conformité des marchandises':                       'BON DE RÉCEPTION',
  'Étiqueter et conditionner les produits':                         'BON DE RÉCEPTION',
  'Ranger et stocker les marchandises':                             'FICHE DE STOCK',
  // CAP EPC — Rayon
  'Approvisionner les rayons':                                      'BON DE COMMANDE',
  'Implanter et baliser les produits':                              "FICHE D'INVENTAIRE",
  "Maintenir la propreté et l'ordre du rayon":                      'RELEVÉ DE DÉMARQUE',
  'Gérer les dates limites de consommation':                        'RELEVÉ DE DÉMARQUE',
  // CAP EPC — Vente
  'Accueillir et orienter les clients':                             'FICHE DÉCOUVERTE CLIENT',
  'Informer et conseiller les clients':                             'FICHE DÉCOUVERTE CLIENT',
  'Réaliser les opérations de caisse':                              'FICHE DE CAISSE',
  'Traiter les réclamations courantes':                             'FICHE DE RÉCLAMATION',
};

// ── Labels ────────────────────────────────────────────────────────────────
const BLOCK_LABELS = {
  theory:      'Apport théorique',
  application: 'Application pratique',
  synthesis:   'Synthèse',
  progressive: 'Exercices progressifs',
  casestudy:   'Étude de cas',
  evaluation:  'Évaluation',
};

// ── Instructions par bloc ─────────────────────────────────────────────────
const BLOCK_INSTRUCTIONS = {
  theory: `APPORT THÉORIQUE (partie numérotée) — champs obligatoires :
- "title" : titre de la partie (ex. "La consommation théorique des denrées")
- "content" : paragraphe d'accroche selon la structure pédagogique choisie :
    • DÉDUCTIF → énonce d'abord la règle/définition, puis illustre avec un exemple concret
    • INDUCTIF  → part d'une situation professionnelle réelle, pose la question, fait observer, puis dégage la règle
- "definition" : "Terme clé : définition complète en une phrase." (null si aucun terme clé)
- "formula"    : formule isolée (ex. "Résultat = A + B − C") — null si aucune formule
- "tables"     : OBLIGATOIRE dès que le contenu comporte des données comparatives (types/définitions/sources, causes/effets, colonnes actif-passif, charges-produits, etc.)
    Format : [{"caption":"Titre optionnel","headers":["Col1","Col2"],"rows":[["v1","v2"]]}]
    Null seulement s'il n'existe vraiment aucun tableau pertinent.
- "example"    : exemple numérique situé dans un établissement fictif français, calcul détaillé étape par étape
- "toRetain"   : "★ À RETENIR — une seule phrase sur la règle essentielle ou le piège à éviter"`,

  application: `APPLICATION PRATIQUE — champs obligatoires :
- "title"          : "Application [lettre]" (A, B, C…)
- "restaurantName" : nom d'entreprise fictive du secteur commercial (supérette, hypermarché, enseigne de distribution, commerce de détail…)
- "documentType"   : choisir le formulaire professionnel adapté à la compétence (BON DE LIVRAISON pour réceptionner, FICHE D'INVENTAIRE pour inventaire, BON DE COMMANDE pour commander, etc.)
- "docRows"        : 2 à 4 lignes de données réalistes pré-remplies dans le formulaire, format selon documentType :
    BON DE COMMANDE   → [["REF","Désignation","Qté","PU HT","TVA%","Total TTC"], ...]
    BON DE LIVRAISON  → [["REF","Désignation","Qté commandée","Qté livrée","Observations"], ...]
    BON DE RÉCEPTION  → [["REF","Désignation","Qté attendue","Qté reçue","État","Observations"], ...]
    FICHE D'INVENTAIRE → [["REF","Désignation","Unité","Stock théo.","Stock réel","Écart","Obs."], ...]
    FICHE DE STOCK    → [["Date","N° doc","Désignation","Entrées","Sorties","Stock"], ...]
    RELEVÉ DE DÉMARQUE     → [["REF","Désignation","DLC/DLUO","Qté","PU","Montant","Cause"], ...]
    FICHE DE CAISSE        → [["Heure","N° ticket","Libellé","Mode paiement","Montant","Rendu monnaie"], ...]
    BON DE COMMANDE CLIENT → [["REF","Désignation","Qté","PU TTC","Remise %","Total TTC"], ...]
    FICHE DÉCOUVERTE CLIENT / FICHE CLÔTURE CAISSE / FICHE DE RÉCLAMATION → null (formulaire auto-généré)
- "context"        : scénario professionnel décrivant la situation (entreprise, date, acteurs). NE PAS répéter la consigne. Données chiffrées utiles à la compréhension.
    CAP EPC → situations simples, vocabulaire de base, données peu nombreuses
    Bac Pro MDCD → situations complexes, données multiples, analyse professionnelle attendue
- "questions"      : 3 objets {q, answer} en progression compréhension → calcul → analyse/conclusion
    Chaque "answer" commence par "Réponse : " et montre le calcul complet étape par étape
- Tous les autres champs : null`,

  synthesis: `SYNTHÈSE DU CHAPITRE — champs obligatoires :
- "content" : phrase d'introduction récapitulative
- "points"  : un objet {number, title, content, formula} par concept clé abordé dans le cours
    "formula" contient la formule isolée si applicable, sinon null
- "tables"  : OBLIGATOIRE — un tableau récapitulatif de TOUS les concepts/formules du chapitre
    Format : [{"caption":"Tableau récapitulatif","headers":["Concept","Définition","Formule"],"rows":[...]}]
- Tous les autres champs : null`,

  progressive: `EXERCICES — exactement 2 exercices complets :
- "exercises" : [{number, restaurantName, context, questions:[{q,answer}], tables}]
    • Exercice 1 : application directe d'un concept (3–4 questions, calculs simples)
    • Exercice 2 : cas intégrant plusieurs concepts + jugement professionnel (4–5 questions)
    Chaque exercice doit avoir un établissement fictif différent.
    Pour chaque exercice, renseigner "documentType" et "docRows" selon le même format que pour APPLICATION :
      "documentType" : formulaire adapté à l'exercice
      "docRows"      : 2 à 4 lignes pré-remplies (colonnes selon documentType) ou null pour FICHE DÉCOUVERTE CLIENT / FICHE CLÔTURE CAISSE / FICHE DE RÉCLAMATION
    "context" : scénario (sans répéter la consigne). NE PAS mettre "Complétez le document..."
    CAP EPC → questions très guidées, calculs simples, vocabulaire accessible
    Bac Pro MDCD → questions plus complexes, données multiples, raisonnement professionnel attendu
    Chaque "answer" commence par "Réponse : " avec calcul complet étape par étape.
- Tous les autres champs : null`,

  casestudy: `ÉTUDE DE CAS — champs obligatoires :
- "restaurantName" : nom d'établissement fictif français
- "context"        : scénario professionnel avec problème de gestion et données numériques
- "questions"      : [{q, answer}] — 1 factuelle, 1–2 calculs, 1 analytique ("Que conseillez-vous ?")
- Autres champs : null`,

  evaluation: `ÉVALUATION FORMATIVE — champs obligatoires :
- "questions" : [{q, answer}] — 2 définitionnelles, 2 de calcul avec données réelles, 1 analytique
- Autres champs : null`,
};

// ── Schéma JSON transmis au modèle ────────────────────────────────────────
const JSON_SCHEMA = `{
  "title": "Titre du cours",
  "subtitle": "Thèmes abordés séparés par ·",
  "companyName": "Nom de l'entreprise fictive utilisée dans TOUT le cours (cohérente avec l'activité : supermarché/hypermarché pour CAP EPC, enseigne de distribution pour Bac Pro MDCD)",
  "objectives": ["objectif 1", "objectif 2", "objectif 3"],
  "sections": [
    {
      "type": "theory|application|synthesis|progressive|casestudy|evaluation",
      "title": "Titre",
      "content": "texte ou null",
      "definition": "Terme : définition ou null",
      "formula": "Formule = A + B ou null",
      "tables": [{"caption":"Titre optionnel","headers":["Col1","Col2"],"rows":[["v1","v2"]]}],
      "example": "Exemple chiffré étape par étape ou null",
      "toRetain": "★ À RETENIR — phrase clé ou null",
      "restaurantName": "Même que companyName ou null",
      "documentType": "BON DE COMMANDE | BON DE LIVRAISON | BON DE RÉCEPTION | FICHE D'INVENTAIRE | FICHE DE STOCK | RELEVÉ DE DÉMARQUE | FICHE DE CAISSE | BON DE COMMANDE CLIENT | FICHE DÉCOUVERTE CLIENT | FICHE CLÔTURE CAISSE | FICHE DE RÉCLAMATION (null pour theory/synthesis)",
      "docRows": [["col1","col2","col3"]],
      "context": "Scénario professionnel avec contexte chiffré (sans répéter la consigne) ou null",
      "questions": [{"q":"Question ?","answer":"Réponse : calcul complet..."}],
      "points": [{"number":1,"title":"Concept","content":"Résumé","formula":"Formule ou null"}],
      "exercises": [{"number":1,"restaurantName":"Même que companyName","documentType":"Type du formulaire","docRows":[["col1","col2"]],"context":"Scénario ou null","tables":null,"questions":[{"q":"...","answer":"Réponse : ..."}]}]
    }
  ]
}`;

// ── Cache temporaire des cours (download par GET) ────────────────────────
const courseCache = new Map();
function storeCourse(course) {
  const id = Math.random().toString(36).slice(2, 10);
  courseCache.set(id, { course, ts: Date.now() });
  // Nettoyage des entrées > 2 h
  for (const [k, v] of courseCache)
    if (Date.now() - v.ts > 7200000) courseCache.delete(k);
  return id;
}

// ── Réparation JSON tronqué ───────────────────────────────────────────────
function extractJson(raw) {
  const str = raw.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '').trim();
  try { return JSON.parse(str); } catch {}
  const arrStart    = str.indexOf('"sections"');
  if (arrStart === -1) throw new Error('Champ "sections" introuvable.');
  const bracketStart = str.indexOf('[', arrStart);
  if (bracketStart === -1) throw new Error('Tableau sections introuvable.');
  let depth = 0, lastEnd = -1, inStr = false, esc = false;
  for (let i = bracketStart + 1; i < str.length; i++) {
    const ch = str[i];
    if (esc)              { esc = false; continue; }
    if (ch === '\\' && inStr) { esc = true; continue; }
    if (ch === '"')       { inStr = !inStr; continue; }
    if (inStr)            { continue; }
    if (ch === '{')       depth++;
    else if (ch === '}')  { depth--; if (depth === 0) lastEnd = i; }
  }
  if (lastEnd === -1) throw new Error('Aucune section JSON complète.');
  try { return JSON.parse(str.slice(0, lastEnd + 1) + ']}'); } catch (e) {
    throw new Error('Impossible de réparer le JSON : ' + e.message);
  }
}

// ── Génération ────────────────────────────────────────────────────────────
app.post('/api/generate', async (req, res) => {
  const { diplome, activite, competence, classe, duree, pedagogyStructure, customInstructions } = req.body;

  if (!diplome)    return res.status(400).json({ error: 'Le diplôme est requis.' });
  if (!activite)   return res.status(400).json({ error: "L'activité est requise." });
  if (!competence) return res.status(400).json({ error: 'La compétence est requise.' });
  if (!classe)     return res.status(400).json({ error: 'La classe est requise.' });

  // Blocs : priorité au tableau envoyé par le frontend, sinon déduit de la durée
  const blocks = Array.isArray(req.body.blocks) && req.body.blocks.length
    ? req.body.blocks
    : duree === '2h'
      ? ['theory', 'application', 'theory', 'application', 'synthesis', 'progressive']
      : ['theory', 'application', 'synthesis'];

  const refData      = REFERENTIELS[diplome] || {};
  const activiteData = refData.activites?.[activite] || {};
  const diplomeLabel = refData.label || diplome;
  const activiteLabel = activiteData.label || activite;

  const isInductive = pedagogyStructure === 'inductive';

  const classeLevel = {
    '2nde Bac Pro':    'beginner',
    '1ère Bac Pro':    'intermediate',
    'Terminale Bac Pro': 'advanced',
    'CAP 1ère année':  'beginner',
    'CAP 2ème année':  'intermediate',
  }[classe] || 'intermediate';

  const levelGuidance = {
    beginner:     'Vocabulaire simple, définitions explicites, analogies du quotidien. Explications très progressives, aucun prérequis supposé.',
    intermediate: 'Vocabulaire professionnel introduit et défini. Exemples ancrés dans la réalité du secteur commercial et de la distribution.',
    advanced:     'Vocabulaire technique assumé. Cas complexes, nuances réglementaires, comparaisons de méthodes et de supports professionnels.',
  }[classeLevel];

  const pedagogyGuidance = isInductive
    ? `STRUCTURE INDUCTIVE : pour chaque apport théorique, commence par une situation professionnelle concrète en commerce/distribution
("Vous observez que…", "Ce matin, à la réception d'une livraison, vous constatez…"),
fais analyser et observer, puis dégage la règle, la définition ou la procédure en conclusion.`
    : `STRUCTURE DÉDUCTIVE : pour chaque apport théorique, énonce d'abord la règle/définition/procédure officielle,
puis illustre immédiatement avec un exemple numérique concret situé dans un commerce ou une grande surface fictive.`;

  const sectionsList = blocks.map((type, i) =>
    `Section ${i + 1} [${BLOCK_LABELS[type] || type}] :\n${BLOCK_INSTRUCTIONS[type] || type}`
  ).join('\n\n');

  const prompt = `Tu es un formateur expert en enseignement professionnel pour les diplômes du commerce et de la distribution.

CONTEXTE DE LA FORMATION :
• Diplôme : ${diplomeLabel}
• Activité du référentiel officiel : ${activiteLabel}
• Compétence visée : ${competence}
• Classe : ${classe}
• Durée de la séance : ${duree || '1h'}
• Approche pédagogique : ${isInductive ? "INDUCTIVE (partir de l'observation vers la règle)" : "DÉDUCTIVE (partir de la règle vers l'application)"}

ADAPTATION AU NIVEAU — ${classe} :
${levelGuidance}

${pedagogyGuidance}

SUJET DU COURS : "${competence}"
dans le cadre de l'activité "${activiteLabel}" — ${diplomeLabel}

STYLE OBLIGATOIRE (reproduire le style des cours professionnels français) :
• Choisir UNE SEULE entreprise fictive cohérente (ex. "Supermarché VALDIS" pour CAP EPC, "MDIS Distribution" pour Bac Pro MDCD) et l'utiliser dans TOUT le cours — placer ce nom dans "companyName" et dans chaque "restaurantName"
• Pour chaque application et exercice, utiliser le formulaire professionnel adapté à la compétence et le placer dans "documentType". Pour la compétence "${competence}", le document recommandé est "${COMPETENCE_DOCUMENTS[competence] || 'le formulaire le plus adapté'}"
• Établissements fictifs français dans TOUS les exemples : magasins, supermarchés, hypermarchés, commerces de détail, entrepôts logistiques, entreprises de distribution
• Calculs TOUJOURS détaillés étape par étape : "320 × 0,250 kg = 80 kg → 80 × 18 € = 1 440 €"
• Définitions isolées sur leur ligne : "Terme : définition complète."
• Formules isolées sur leur ligne : "Résultat = A + B − C"
• Tableaux obligatoires dès que le contenu comporte des données comparatives (documents commerciaux, procédures, grilles de contrôle…)
• ★ À RETENIR en fin de chaque apport théorique : une seule phrase essentielle
• Réponses des applications : "Réponse : calcul complet" en italique
• Vocabulaire et situations professionnelles du commerce et de la distribution
• Ton professionnel, rigoureux, conforme aux exigences du référentiel
${customInstructions ? `\nCONSIGNES DE L'ENSEIGNANT (prioritaires) :\n${customInstructions}` : ''}

STRUCTURE DU COURS (${blocks.length} sections — respecte l'ordre et le contenu de chaque section) :
${sectionsList}

Réponds UNIQUEMENT avec un objet JSON valide (aucun markdown, aucune balise code) :
${JSON_SCHEMA}`;

  try {
    const response = await fetch(ANTHROPIC_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': API_KEY,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 8192,
        messages: [{ role: 'user', content: prompt }]
      })
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      return res.status(500).json({ error: err.error?.message || `Erreur API (${response.status})` });
    }

    const data = await response.json();
    if (data.stop_reason === 'max_tokens') console.warn('Réponse tronquée — réparation JSON');
    const course     = extractJson(data.content[0].text.trim());
    const downloadId = storeCourse(course);
    res.json({ course, downloadId });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ── Helpers docx ──────────────────────────────────────────────────────────
const run = (text, opts = {}) =>
  new TextRun({ text: String(text ?? ''), size: 22, ...opts });

const body = (runsOrText, paraOpts = {}) =>
  new Paragraph({
    children: typeof runsOrText === 'string'
      ? [run(runsOrText)]
      : Array.isArray(runsOrText) ? runsOrText : [runsOrText],
    spacing: { after: 100 },
    ...paraOpts
  });

const divider = () =>
  new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'E5E7EB' } },
    spacing: { before: 240, after: 240 }
  });

const callout = (runsOrText, bgColor, borderColor, extraOpts = {}) =>
  new Paragraph({
    children: typeof runsOrText === 'string'
      ? [run(runsOrText)]
      : Array.isArray(runsOrText) ? runsOrText : [runsOrText],
    shading: { type: ShadingType.CLEAR, fill: bgColor },
    border: { left: { style: BorderStyle.THICK, size: 18, color: borderColor, space: 8 } },
    indent: { left: 200, right: 160 },
    spacing: { before: 80, after: 80 },
    ...extraOpts
  });

const typeHeader = (label, bgColor, textColor) =>
  new Paragraph({
    children: [run(`     ${label.toUpperCase()}     `, { bold: true, size: 22, color: textColor })],
    alignment: AlignmentType.CENTER,
    shading: { type: ShadingType.CLEAR, fill: bgColor },
    spacing: { before: 280, after: 100 }
  });

function makeTable(headers, rows) {
  const border = (color, size = 4) => ({ style: BorderStyle.SINGLE, size, color });

  const hdrRow = new TableRow({
    tableHeader: true,
    children: (headers || []).map(h =>
      new TableCell({
        children: [new Paragraph({
          children: [run(String(h ?? ''), { bold: true, size: 20 })],
          spacing: { before: 60, after: 60 }
        })],
        shading: { type: ShadingType.CLEAR, fill: 'EEF2FF' },
        borders: {
          top:    border('A5B4FC'), bottom: border('A5B4FC'),
          left:   border('A5B4FC'), right:  border('A5B4FC')
        },
        margins: { left: 100, right: 100 }
      })
    )
  });

  const dataRows = (rows || []).map((row, ri) =>
    new TableRow({
      children: (row || []).map(cell =>
        new TableCell({
          children: [new Paragraph({
            children: [run(String(cell ?? ''), { size: 20 })],
            spacing: { before: 60, after: 60 }
          })],
          shading: { type: ShadingType.CLEAR, fill: ri % 2 === 0 ? 'F9FAFB' : 'FFFFFF' },
          borders: {
            top:    border('D1D5DB', 2), bottom: border('D1D5DB', 2),
            left:   border('D1D5DB', 2), right:  border('D1D5DB', 2)
          },
          margins: { left: 100, right: 100 }
        })
      )
    })
  );

  return new Table({
    rows: [hdrRow, ...dataRows],
    width: { size: 100, type: WidthType.PERCENTAGE }
  });
}

// ── En-tête de formulaire professionnel (fond gris D9D9D9, Calibri 10pt) ─
function makeFormHeader(companyName, documentType) {
  const gb = 'D9D9D9';
  const bd = () => ({ style: BorderStyle.SINGLE, size: 4, color: '999999' });
  const borders = { top: bd(), bottom: bd(), left: bd(), right: bd() };
  const mar = { left: 80, right: 80 };

  return new Table({
    rows: [new TableRow({ children: [
      new TableCell({
        children: [
          new Paragraph({ children: [run(companyName || '..........', { bold: true, size: 20, font: 'Calibri' })], spacing: { before: 40, after: 20 } }),
          new Paragraph({ children: [run('Adresse : ..........', { size: 18, font: 'Calibri', color: '6B7280' })], spacing: { before: 0, after: 40 } }),
        ],
        shading: { type: ShadingType.CLEAR, fill: gb }, borders, margins: mar,
        width: { size: 40, type: WidthType.PERCENTAGE }
      }),
      new TableCell({
        children: [new Paragraph({ children: [run(documentType || 'DOCUMENT', { bold: true, size: 24, font: 'Calibri', allCaps: true })], alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 } })],
        shading: { type: ShadingType.CLEAR, fill: gb }, borders, margins: mar,
        width: { size: 35, type: WidthType.PERCENTAGE }
      }),
      new TableCell({
        children: [
          new Paragraph({ children: [run('Réf : ..........', { size: 18, font: 'Calibri' })], spacing: { before: 20, after: 10 } }),
          new Paragraph({ children: [run('N° : ..........', { size: 18, font: 'Calibri' })], spacing: { before: 0, after: 10 } }),
          new Paragraph({ children: [run('Date : ..........', { size: 18, font: 'Calibri' })], spacing: { before: 0, after: 20 } }),
        ],
        shading: { type: ShadingType.CLEAR, fill: gb }, borders, margins: mar,
        width: { size: 25, type: WidthType.PERCENTAGE }
      }),
    ]})],
    width: { size: 100, type: WidthType.PERCENTAGE }
  });
}

// ── Pied de formulaire professionnel ─────────────────────────────────────
function makeFormFooter() {
  const bd = () => ({ style: BorderStyle.SINGLE, size: 4, color: '999999' });
  const borders = { top: bd(), bottom: bd(), left: bd(), right: bd() };
  const mar = { left: 80, right: 80 };
  const cell = (lines) => new TableCell({
    children: lines.map(l => new Paragraph({ children: [run(l, { size: 18, font: 'Calibri' })], spacing: { before: 28, after: 28 } })),
    borders, margins: mar, width: { size: 50, type: WidthType.PERCENTAGE }
  });
  return new Table({
    rows: [new TableRow({ children: [
      cell(['Émis par : ..........', 'Date : ..........', 'Service : ..........']),
      cell(['Vérifié par : ..........', 'Signature : ..........', 'Tampon entreprise :']),
    ]})],
    width: { size: 100, type: WidthType.PERCENTAGE }
  });
}

// ── Templates de documents professionnels ────────────────────────────────
const docR = (txt, opts = {}) =>
  new TextRun({ text: String(txt ?? ''), size: 20, font: 'Calibri', ...opts });
const docP = (content, pOpts = {}) =>
  new Paragraph({
    children: typeof content === 'string' ? [docR(content)]
              : Array.isArray(content) ? content : [content],
    spacing: { before: 36, after: 36 }, ...pOpts,
  });
const fb    = () => ({ style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' });
const fbAll = () => { const b = fb(); return { top:b, bottom:b, left:b, right:b }; };
const fCell = (children, wPct, fill, span) => new TableCell({
  children: Array.isArray(children) ? children : [children],
  borders: fbAll(), margins: { left: 80, right: 80 },
  ...(wPct ? { width: { size: wPct, type: WidthType.PERCENTAGE } } : {}),
  ...(fill ? { shading: { type: ShadingType.CLEAR, fill } } : {}),
  ...(span ? { columnSpan: span } : {}),
});

function templateHeader(company, title) {
  return new Table({
    rows: [new TableRow({ children: [
      fCell([
        docP([docR(company || 'ENTREPRISE', { bold: true, size: 24 })]),
        docP([docR('Adresse : …………………………………………')]),
        docP([docR('Tél : ………………………  Email : ……………………')]),
        docP([docR('SIRET : …………………………………………………')]),
      ], 55, 'D9D9D9'),
      fCell([
        docP([docR(title || 'DOCUMENT', { bold: true, size: 28, allCaps: true })], { alignment: AlignmentType.CENTER }),
        docP([docR('N° : ___________')], { alignment: AlignmentType.RIGHT }),
        docP([docR('Date : _________')], { alignment: AlignmentType.RIGHT }),
        docP([docR('Réf : __________')], { alignment: AlignmentType.RIGHT }),
      ], 45, 'D9D9D9'),
    ]})],
    width: { size: 100, type: WidthType.PERCENTAGE }
  });
}

function templateFooter() {
  const col = (lbl) => fCell([
    docP([docR(lbl, { bold: true })]),
    docP(''),
    docP([docR('Date :      _______________')]),
    docP([docR('Signature : _______________')]),
  ], 34);
  return new Table({
    rows: [new TableRow({ children: [col('Établi par :'), col('Vérifié par :'), col('Approuvé par :')] })],
    width: { size: 100, type: WidthType.PERCENTAGE }
  });
}

function docDataTable(cols, dataRows, totalEmpty) {
  const filled = (dataRows || []).filter(r => r?.length);
  const empty  = Math.max(0, (totalEmpty || 5) - filled.length);
  const hdr = new TableRow({ tableHeader: true, children: cols.map(c =>
    fCell([docP([docR(c.label, { bold: true })], { alignment: AlignmentType.CENTER })], c.w, 'D9D9D9')
  )});
  return new Table({
    rows: [
      hdr,
      ...filled.map(row => new TableRow({ children: cols.map((c, ci) => fCell([docP([docR(row[ci] ?? '')])], c.w)) })),
      ...Array.from({ length: empty }, () => new TableRow({
        height: { value: 340, rule: 'atLeast' },
        children: cols.map(c => fCell([docP('')], c.w))
      })),
    ],
    width: { size: 100, type: WidthType.PERCENTAGE }
  });
}

function totalsTable(lines) {
  return new Table({
    rows: lines.map(([lbl, val]) => new TableRow({ children: [
      fCell([docP([docR(lbl, { bold: true })], { alignment: AlignmentType.RIGHT })], 75, 'F3F4F6'),
      fCell([docP([docR(val || '')])], 25),
    ]})),
    width: { size: 100, type: WidthType.PERCENTAGE }
  });
}

function infoRow(left, right) {
  return new Table({
    rows: [new TableRow({ children: [
      fCell(left.map(l => docP([docR(l)])), 50),
      fCell(right.map(r => docP([docR(r)])), 50),
    ]})],
    width: { size: 100, type: WidthType.PERCENTAGE }
  });
}

function buildBonCommande(company, rows) {
  return [
    templateHeader(company, 'BON DE COMMANDE'),
    new Paragraph({ spacing: { before: 60 } }),
    infoRow(
      ['Fournisseur : ………………………………………………', 'Adresse : …………………………………………………', 'Contact : ……………………………………………………'],
      ['Livraison prévue : __________________', 'Modalités paiement : _______________', 'Conditions : _______________________']
    ),
    new Paragraph({ spacing: { before: 60 } }),
    docDataTable([
      { label:'Réf', w:9 }, { label:'Désignation', w:33 }, { label:'Qté', w:9 },
      { label:'PU HT', w:16 }, { label:'TVA %', w:12 }, { label:'Total TTC', w:21 }
    ], rows, 6),
    new Paragraph({ spacing: { before: 40 } }),
    totalsTable([['Total HT :', ''], ['TVA :', ''], ['Total TTC :', '']]),
    new Paragraph({ spacing: { before: 80 } }),
    templateFooter(),
    new Paragraph({ spacing: { after: 100 } }),
  ];
}

function buildBonLivraison(company, rows) {
  return [
    templateHeader(company, 'BON DE LIVRAISON'),
    new Paragraph({ spacing: { before: 60 } }),
    infoRow(
      ['Expéditeur : …………………………………………………', 'Transporteur : ………………………………………'],
      ['Destinataire : ………………………………………………', 'Heure livraison : __:__']
    ),
    new Paragraph({ spacing: { before: 60 } }),
    docDataTable([
      { label:'Réf', w:10 }, { label:'Désignation', w:32 },
      { label:'Qté commandée', w:18 }, { label:'Qté livrée', w:18 }, { label:'Observations', w:22 }
    ], rows, 7),
    new Paragraph({ spacing: { before: 80 } }),
    templateFooter(),
    new Paragraph({ spacing: { after: 100 } }),
  ];
}

function buildBonReception(company, rows) {
  return [
    templateHeader(company, 'BON DE RÉCEPTION'),
    new Paragraph({ spacing: { before: 60 } }),
    infoRow(
      ['Fournisseur : …………………………………………………', 'N° BL fournisseur : _______________'],
      ['Date réception : ___________________', 'Réceptionnaire : __________________']
    ),
    new Paragraph({ spacing: { before: 60 } }),
    docDataTable([
      { label:'Réf', w:9 }, { label:'Désignation', w:30 }, { label:'Qté attendue', w:14 },
      { label:'Qté reçue', w:14 }, { label:'État (C/NC)', w:13 }, { label:'Observations', w:20 }
    ], rows, 6),
    new Paragraph({ spacing: { before: 80 } }),
    templateFooter(),
    new Paragraph({ spacing: { after: 100 } }),
  ];
}

function buildFicheInventaire(company, rows) {
  return [
    templateHeader(company, "FICHE D'INVENTAIRE"),
    new Paragraph({ spacing: { before: 60 } }),
    infoRow(
      ['Date inventaire : ___________________', 'Rayon / Zone : _____________________'],
      ['Responsable : ______________________', 'Méthode :  □ Comptage   □ Scan code-barres']
    ),
    new Paragraph({ spacing: { before: 60 } }),
    docDataTable([
      { label:'Réf', w:9 }, { label:'Désignation', w:28 }, { label:'Unité', w:8 },
      { label:'Stock théo.', w:14 }, { label:'Stock réel', w:14 }, { label:'Écart', w:13 }, { label:'Observations', w:14 }
    ], rows, 7),
    new Paragraph({ spacing: { before: 40 } }),
    totalsTable([['Valeur totale du stock :', ''], ['Écart de valeur (€) :', ''], ['Observations générales :', '']]),
    new Paragraph({ spacing: { before: 80 } }),
    templateFooter(),
    new Paragraph({ spacing: { after: 100 } }),
  ];
}

function buildFicheStock(company, rows) {
  return [
    templateHeader(company, 'FICHE DE STOCK'),
    new Paragraph({ spacing: { before: 60 } }),
    infoRow(
      ['Produit : ………………………………………………………', 'Référence : ________________________'],
      ['Unité de mesure : __________________', 'Stock minimum : ____________________']
    ),
    new Paragraph({ spacing: { before: 60 } }),
    docDataTable([
      { label:'Date', w:13 }, { label:'N° doc', w:11 }, { label:'Désignation', w:30 },
      { label:'Entrées', w:15 }, { label:'Sorties', w:15 }, { label:'Stock', w:16 }
    ], rows, 8),
    new Paragraph({ spacing: { before: 80 } }),
    templateFooter(),
    new Paragraph({ spacing: { after: 100 } }),
  ];
}

function buildFicheClôtureCaisse(company) {
  const coupures = ['50 €','20 €','10 €','5 €','2 €','1 €','0,50 €','0,20 €','0,10 €','0,05 €'];
  const sectionA = new Table({
    rows: [
      new TableRow({ children: [new TableCell({
        children: [docP([docR('FOND DE CAISSE — DÉTAIL COUPURES', { bold: true })], { alignment: AlignmentType.CENTER })],
        columnSpan: 3, borders: fbAll(), margins: { left: 80, right: 80 }, shading: { type: ShadingType.CLEAR, fill: 'D9D9D9' }
      })]}),
      new TableRow({ tableHeader: true, children: [
        fCell([docP([docR('Coupure', { bold: true })],          { alignment: AlignmentType.CENTER })], 34, 'D9D9D9'),
        fCell([docP([docR('Nb billets / pièces', { bold: true })], { alignment: AlignmentType.CENTER })], 33, 'D9D9D9'),
        fCell([docP([docR('Montant (€)', { bold: true })],      { alignment: AlignmentType.CENTER })], 33, 'D9D9D9'),
      ]}),
      ...coupures.map(c => new TableRow({
        height: { value: 300, rule: 'atLeast' },
        children: [fCell([docP(c)], 34), fCell([docP('')], 33), fCell([docP('')], 33)]
      })),
      new TableRow({ children: [
        new TableCell({
          children: [docP([docR('TOTAL FOND DE CAISSE', { bold: true })], { alignment: AlignmentType.RIGHT })],
          columnSpan: 2, borders: fbAll(), margins: { left: 80, right: 80 }, shading: { type: ShadingType.CLEAR, fill: 'F3F4F6' }
        }),
        fCell([docP('')], 33),
      ]}),
    ],
    width: { size: 100, type: WidthType.PERCENTAGE }
  });
  return [
    templateHeader(company, 'FICHE CLÔTURE CAISSE'),
    new Paragraph({ spacing: { before: 60 } }),
    infoRow(
      ['Caisse N° : ________________________', 'Date : ____________________________'],
      ['Heure ouverture : __________________', 'Heure clôture : ___________________']
    ),
    new Paragraph({ spacing: { before: 80 } }),
    sectionA,
    new Paragraph({ spacing: { before: 80 } }),
    totalsTable([
      ['Total espèces collectées :', ''], ['Total CB :', ''],
      ['Total chèques :', ''], ['Total tickets restaurant :', ''],
      ['TOTAL GÉNÉRAL :', ''], ['Montant théorique :', ''], ['ÉCART :', ''],
    ]),
    new Paragraph({ spacing: { before: 80 } }),
    templateFooter(),
    new Paragraph({ spacing: { after: 100 } }),
  ];
}

function buildFicheReclamation(company) {
  const r2 = (l, r) => new TableRow({ children: [
    fCell([docP([docR(l, { bold: true })]), new Paragraph({ spacing: { before: 100, after: 100 } })], 50),
    fCell([docP([docR(r, { bold: true })]), new Paragraph({ spacing: { before: 100, after: 100 } })], 50),
  ]});
  const r1 = (l) => new TableRow({ children: [new TableCell({
    children: [docP([docR(l, { bold: true })]), new Paragraph({ spacing: { before: 120, after: 120 } })],
    columnSpan: 2, borders: fbAll(), margins: { left: 80, right: 80 }
  })]});
  return [
    templateHeader(company, 'FICHE DE RÉCLAMATION'),
    new Paragraph({ spacing: { before: 60 } }),
    new Table({
      rows: [
        r2('Date :', 'N° réclamation :'),
        r2('Client :', 'Contact / Tél :'),
        r2('Nature :  □ Produit   □ Service   □ Livraison   □ Facturation', 'Urgence :  □ Haute   □ Moyenne   □ Faible'),
        r1('Produit / Service concerné :'),
        r1('Description du problème :'),
        r1('Action demandée par le client :'),
        r1('Suite donnée / Traitement :'),
        r2('Délai de résolution :', 'Statut :  □ Ouvert   □ En cours   □ Clôturé'),
      ],
      width: { size: 100, type: WidthType.PERCENTAGE }
    }),
    new Paragraph({ spacing: { before: 80 } }),
    templateFooter(),
    new Paragraph({ spacing: { after: 100 } }),
  ];
}

function buildRelevéDémarque(company, rows) {
  return [
    templateHeader(company, 'RELEVÉ DE DÉMARQUE'),
    new Paragraph({ spacing: { before: 60 } }),
    infoRow(
      ['Date : _________________________', 'Rayon / Zone : _________________'],
      ['Responsable : __________________', 'Visa responsable : _____________']
    ),
    new Paragraph({ spacing: { before: 60 } }),
    docDataTable([
      { label:'Réf', w:9 }, { label:'Désignation', w:27 }, { label:'DLC/DLUO', w:12 },
      { label:'Qté', w:8 }, { label:'PU (€)', w:10 }, { label:'Montant (€)', w:12 },
      { label:'Cause  □ Périmé  □ Casse  □ Vol  □ Autre', w:22 }
    ], rows, 7),
    new Paragraph({ spacing: { before: 40 } }),
    totalsTable([['Valeur totale démarque (€) :', ''], ['Observations :', '']]),
    new Paragraph({ spacing: { before: 80 } }),
    templateFooter(),
    new Paragraph({ spacing: { after: 100 } }),
  ];
}

function buildFicheCaisse(company, rows) {
  return [
    templateHeader(company, 'FICHE DE CAISSE'),
    new Paragraph({ spacing: { before: 60 } }),
    infoRow(
      ['Caisse N° : ____________________', 'Date : ____________________________'],
      ["Heure d'ouverture : ____________", 'Heure de clôture : ________________']
    ),
    new Paragraph({ spacing: { before: 60 } }),
    docDataTable([
      { label:'Heure', w:12 }, { label:'N° ticket', w:12 }, { label:'Libellé', w:30 },
      { label:'Mode paiement', w:18 }, { label:'Montant (€)', w:14 }, { label:'Rendu monnaie', w:14 }
    ], rows, 8),
    new Paragraph({ spacing: { before: 40 } }),
    totalsTable([
      ['Total transactions :', ''], ['Montant théorique :', ''],
      ['Montant compté :', ''], ['ÉCART :', '']
    ]),
    new Paragraph({ spacing: { before: 80 } }),
    templateFooter(),
    new Paragraph({ spacing: { after: 100 } }),
  ];
}

function buildBonCommandeClient(company, rows) {
  return [
    templateHeader(company, 'BON DE COMMANDE CLIENT'),
    new Paragraph({ spacing: { before: 60 } }),
    infoRow(
      ['Client : ……………………………………………………………', 'Adresse : …………………………………………………', 'Contact / Tél : ……………………………………'],
      ['Date commande : ___________________', 'Livraison souhaitée : ______________', 'Vendeur : _________________________']
    ),
    new Paragraph({ spacing: { before: 60 } }),
    docDataTable([
      { label:'Réf', w:9 }, { label:'Désignation', w:31 }, { label:'Qté', w:9 },
      { label:'PU TTC (€)', w:15 }, { label:'Remise %', w:12 }, { label:'Total TTC (€)', w:24 }
    ], rows, 6),
    new Paragraph({ spacing: { before: 40 } }),
    totalsTable([['Sous-total HT :', ''], ['TVA :', ''], ['Total TTC :', ''], ['Remise accordée :', '']]),
    new Paragraph({ spacing: { before: 80 } }),
    templateFooter(),
    new Paragraph({ spacing: { after: 100 } }),
  ];
}

function buildFicheDecouverteClient(company) {
  const r2 = (l, r) => new TableRow({ children: [
    fCell([docP([docR(l, { bold: true })]), new Paragraph({ spacing: { before: 100, after: 100 } })], 50),
    fCell([docP([docR(r, { bold: true })]), new Paragraph({ spacing: { before: 100, after: 100 } })], 50),
  ]});
  const r1 = (l) => new TableRow({ children: [new TableCell({
    children: [docP([docR(l, { bold: true })]), new Paragraph({ spacing: { before: 120, after: 120 } })],
    columnSpan: 2, borders: fbAll(), margins: { left: 80, right: 80 }
  })]});
  return [
    templateHeader(company, 'FICHE DÉCOUVERTE CLIENT'),
    new Paragraph({ spacing: { before: 60 } }),
    new Table({
      rows: [
        new TableRow({ children: [new TableCell({
          children: [docP([docR('IDENTIFICATION DU PROSPECT', { bold: true })], { alignment: AlignmentType.CENTER })],
          columnSpan: 2, borders: fbAll(), margins: { left: 80, right: 80 }, shading: { type: ShadingType.CLEAR, fill: 'D9D9D9' }
        })]}),
        r2('Nom / Prénom :', 'Société / Raison sociale :'),
        r2('Téléphone :', 'Email :'),
        r2('Adresse :', "Secteur d'activité :"),
        new TableRow({ children: [new TableCell({
          children: [docP([docR('BESOINS ET PROJET', { bold: true })], { alignment: AlignmentType.CENTER })],
          columnSpan: 2, borders: fbAll(), margins: { left: 80, right: 80 }, shading: { type: ShadingType.CLEAR, fill: 'D9D9D9' }
        })]}),
        r1('Description du besoin / projet :'),
        r2('Délai souhaité :', 'Budget estimé (€) :'),
        r1('Produits / Services présentés :'),
        r1('Objections et réponses apportées :'),
        new TableRow({ children: [new TableCell({
          children: [docP([docR('DÉCISION ET SUITES', { bold: true })], { alignment: AlignmentType.CENTER })],
          columnSpan: 2, borders: fbAll(), margins: { left: 80, right: 80 }, shading: { type: ShadingType.CLEAR, fill: 'D9D9D9' }
        })]}),
        r2('Décisionnaire (Nom / Fonction) :', 'Maturité :  □ Chaud   □ Tiède   □ Froid'),
        r2('Date de relance :', 'Statut :  □ Devis à envoyer   □ En négociation   □ Converti   □ Perdu'),
        r1('Commentaires / Observations :'),
      ],
      width: { size: 100, type: WidthType.PERCENTAGE }
    }),
    new Paragraph({ spacing: { before: 80 } }),
    templateFooter(),
    new Paragraph({ spacing: { after: 100 } }),
  ];
}

function buildDocumentTemplate(documentType, company, rows) {
  const t = (documentType || '').toUpperCase();
  if (t.includes('DÉCOUVERTE') || t.includes('DECOUVERTE'))                 return buildFicheDecouverteClient(company);
  if (t.includes('COMMANDE CLIENT'))                                        return buildBonCommandeClient(company, rows);
  if (t.includes('COMMANDE'))                                               return buildBonCommande(company, rows);
  if (t.includes('LIVRAISON'))                                              return buildBonLivraison(company, rows);
  if (t.includes('RÉCEPTION') || t.includes('RECEPTION'))                  return buildBonReception(company, rows);
  if (t.includes('INVENTAIRE'))                                             return buildFicheInventaire(company, rows);
  if (t.includes('DÉMARQUE') || t.includes('DEMARQUE'))                    return buildRelevéDémarque(company, rows);
  if (t.includes('CLÔTURE') || t.includes('CLOTURE'))                      return buildFicheClôtureCaisse(company);
  if (t.includes('CAISSE'))                                                 return buildFicheCaisse(company, rows);
  if (t.includes('STOCK'))                                                  return buildFicheStock(company, rows);
  if (t.includes('RÉCLAMATION') || t.includes('RECLAMATION'))              return buildFicheReclamation(company);
  return [templateHeader(company, documentType || 'DOCUMENT'), new Paragraph({ spacing: { before: 100, after: 200 } }), templateFooter(), new Paragraph({ spacing: { after: 100 } })];
}

// ── Fonction de génération du buffer docx ────────────────────────────────
async function buildDocx(course) {

  try {
    const children = [];
    let partieNum = 0;

    // ── Titre ──
    children.push(new Paragraph({
      children: [run(course.title, { bold: true, size: 52, color: '1A1A1A' })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 }
    }));
    if (course.subtitle) {
      children.push(new Paragraph({
        children: [run(course.subtitle, { italics: true, size: 20, color: '6B7280' })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 320 }
      }));
    }

    // ── Objectifs ──
    if (course.objectives?.length) {
      children.push(divider());
      children.push(new Paragraph({
        text: 'Objectifs pédagogiques',
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 100, after: 100 }
      }));
      course.objectives.forEach(o => children.push(body(`• ${o}`)));
    }

    // ── Sections ──
    course.sections.forEach(sec => {
      children.push(divider());

      switch (sec.type) {

        // ── APPORT THÉORIQUE ──────────────────────────────────────────────
        case 'theory': {
          partieNum++;
          children.push(new Paragraph({
            children: [
              run(`PARTIE ${partieNum}  `, { bold: true, size: 28, allCaps: true }),
              run(sec.title, { bold: true, size: 28 })
            ],
            spacing: { before: 160, after: 100 }
          }));

          children.push(typeHeader('Apport théorique', 'EEF2FF', '4338CA'));

          if (sec.content) {
            sec.content.split('\n').filter(l => l.trim()).forEach(l =>
              children.push(body(l))
            );
          }

          if (sec.definition) {
            const ci    = sec.definition.indexOf(':');
            const term  = ci > -1 ? sec.definition.slice(0, ci + 1) : '';
            const def   = ci > -1 ? sec.definition.slice(ci + 1)    : sec.definition;
            children.push(callout(
              term ? [run(term, { bold: true }), run(def)] : [run(def)],
              'EEF2FF', '6366F1'
            ));
          }

          if (sec.formula) {
            children.push(callout(
              [run(sec.formula, { bold: true, size: 24, color: '166534' })],
              'F0FDF4', '22C55E',
              { alignment: AlignmentType.CENTER }
            ));
          }

          if (sec.tables?.length) {
            sec.tables.forEach(t => {
              if (t.caption) children.push(body(run(t.caption, { italics: true, color: '6B7280', size: 20 })));
              children.push(new Paragraph({ spacing: { before: 80 } }));
              children.push(makeTable(t.headers, t.rows));
              children.push(new Paragraph({ spacing: { after: 80 } }));
            });
          }

          if (sec.example) {
            children.push(callout(
              [run('Exemple : ', { bold: true }), run(sec.example, { italics: true })],
              'F9FAFB', 'D1D5DB'
            ));
          }

          if (sec.toRetain) {
            children.push(callout(
              [run(sec.toRetain, { bold: true, color: '78350F' })],
              'FFFBEB', 'F59E0B'
            ));
          }
          break;
        }

        // ── APPLICATION ───────────────────────────────────────────────────
        case 'application':
        case 'casestudy': {
          const appLabel = sec.type === 'application' ? 'Application' : 'Étude de cas';
          children.push(typeHeader(appLabel, 'F0FDF4', '15803D'));

          if (sec.restaurantName) {
            children.push(new Paragraph({
              children: [
                run(sec.title || appLabel, { bold: true, size: 24, color: '1D4ED8' }),
                run(`    ${sec.restaurantName}`, { bold: true, size: 22, italics: true, color: '374151' })
              ],
              border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'BFDBFE' } },
              spacing: { before: 120, after: 100 }
            }));
          }

          // Consigne
          children.push(body(
            run('Complétez le document suivant à partir des informations ci-dessous :', { bold: true, color: '1D4ED8' }),
            { spacing: { before: 80, after: 80 } }
          ));
          // Scénario
          if (sec.context) {
            children.push(callout([run(sec.context, { italics: true, color: '374151' })], 'F9FAFB', 'D1D5DB'));
            children.push(new Paragraph({ spacing: { before: 80 } }));
          }
          // Template de document professionnel
          buildDocumentTemplate(sec.documentType, course.companyName || sec.restaurantName, sec.docRows)
            .forEach(el => children.push(el));

          (sec.questions || []).forEach((qa, qi) => {
            children.push(body(run(`${qi + 1}.  ${qa.q}`, { bold: true }),
              { spacing: { before: 120, after: 40 } }));
            if (qa.answer) {
              children.push(body(run(qa.answer, { italics: true, color: 'DC2626' }),
                { indent: { left: 300 }, spacing: { before: 0, after: 80 } }));
            }
          });
          break;
        }

        // ── SYNTHÈSE ──────────────────────────────────────────────────────
        case 'synthesis': {
          children.push(new Paragraph({
            text: 'Synthèse du chapitre',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 160, after: 100 }
          }));
          if (sec.content) children.push(body(sec.content));

          (sec.points || []).forEach(pt => {
            children.push(new Paragraph({
              children: [
                run(`${pt.number} · `, { bold: true, size: 26, color: '5B4FCF' }),
                run(String(pt.title ?? ''), { bold: true, size: 26, color: '1A1A1A' })
              ],
              spacing: { before: 200, after: 60 }
            }));
            if (pt.content) children.push(body(pt.content));
            if (pt.formula) {
              children.push(callout(
                [run(pt.formula, { bold: true, color: '166534' })],
                'F0FDF4', '22C55E'
              ));
            }
          });
          break;
        }

        // ── EXERCICES PROGRESSIFS ─────────────────────────────────────────
        case 'progressive': {
          children.push(new Paragraph({
            text: 'Exercices',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 160, after: 100 }
          }));

          (sec.exercises || []).forEach(ex => {
            children.push(new Paragraph({
              children: [
                run(`Exercice ${ex.number}  `, { bold: true, size: 24, color: 'B45309' }),
                run(String(ex.restaurantName ?? ''), { size: 22, italics: true, color: '374151' })
              ],
              shading: { type: ShadingType.CLEAR, fill: 'FFFBEB' },
              border: {
                top:    { style: BorderStyle.SINGLE, size: 4, color: 'FDE68A' },
                bottom: { style: BorderStyle.SINGLE, size: 4, color: 'FDE68A' }
              },
              spacing: { before: 240, after: 80 }
            }));

            // Consigne
            children.push(body(
              run('Complétez le document suivant à partir des informations ci-dessous :', { bold: true, color: '1D4ED8' }),
              { spacing: { before: 80, after: 80 } }
            ));
            // Scénario
            if (ex.context) {
              children.push(callout([run(ex.context, { italics: true, color: '374151' })], 'F9FAFB', 'D1D5DB'));
              children.push(new Paragraph({ spacing: { before: 80 } }));
            }
            // Template de document professionnel
            buildDocumentTemplate(ex.documentType, course.companyName || ex.restaurantName, ex.docRows)
              .forEach(el => children.push(el));

            (ex.questions || []).forEach((qa, qi) => {
              children.push(body(run(`${qi + 1}.  ${qa.q}`, { bold: true }),
                { spacing: { before: 120, after: 40 } }));
              if (qa.answer) {
                children.push(body(run(qa.answer, { italics: true, color: 'DC2626' }),
                  { indent: { left: 300 }, spacing: { before: 0, after: 80 } }));
              }
            });
          });
          break;
        }

        // ── ÉVALUATION ────────────────────────────────────────────────────
        case 'evaluation': {
          children.push(new Paragraph({
            text: 'Évaluation',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 160, after: 100 }
          }));
          (sec.questions || []).forEach((qa, qi) => {
            children.push(body(run(`${qi + 1}.  ${qa.q}`, { bold: true }),
              { spacing: { before: 120, after: 40 } }));
            if (qa.answer) {
              children.push(body(run(qa.answer, { italics: true, color: 'DC2626' }),
                { indent: { left: 300 }, spacing: { before: 0, after: 80 } }));
            }
          });
          break;
        }

        default:
          if (sec.content) sec.content.split('\n').filter(l => l.trim()).forEach(l => children.push(body(l)));
          (sec.keyPoints || []).forEach(kp => children.push(body(`• ${kp}`)));
          if (sec.exercise) children.push(callout([run(sec.exercise, { italics: true })], 'FFF7ED', 'F97316'));
      }
    });

    const doc      = new Document({ sections: [{ children }] });
    const buffer   = await Packer.toBuffer(doc);
    const filename = course.title.replace(/[^a-z0-9]/gi, '_').toLowerCase() + '.docx';
    return { buffer, filename };
  } catch (err) { throw err; }
}

// ── Headers communs pour téléchargement iframe-compatible ─────────────────
function setDownloadHeaders(res, filename) {
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('X-Frame-Options', 'ALLOWALL');
  res.setHeader('Content-Security-Policy', 'frame-ancestors *');
}

// ── GET /api/download/:id — ouvert via window.open() depuis l'iframe ──────
app.get('/api/download/:id', async (req, res) => {
  const entry = courseCache.get(req.params.id);
  if (!entry) return res.status(404).send('Lien expiré ou invalide. Régénérez le cours.');
  try {
    const { buffer, filename } = await buildDocx(entry.course);
    setDownloadHeaders(res, filename);
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).send('Erreur génération docx : ' + err.message);
  }
});

// ── POST /api/download — rétrocompatibilité ───────────────────────────────
app.post('/api/download', async (req, res) => {
  const { course } = req.body;
  if (!course) return res.status(400).json({ error: 'Données du cours manquantes.' });
  try {
    const { buffer, filename } = await buildDocx(course);
    setDownloadHeaders(res, filename);
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => console.log(`Serveur démarré → http://localhost:${PORT}`));
