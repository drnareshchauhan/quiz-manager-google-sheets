/*************************************************
 * PSM TOPIC MASTER
 *************************************************/
const PSM_TOPICS = {
  1: "History of Medicine",
  2: "Concepts of Health and Disease",
  3: "Epidemiology and Vaccines",
  4: "Screening of Disease",
  5: "Communicable & Non-communicable Diseases",
  6: "National Health Programmes",
  7: "Demography, Family Planning and Contraception",
  8: "Preventive Obstetrics, Paediatrics and Geriatrics",
  9: "Nutrition",
  10: "Social Sciences",
  11: "Environment",
  12: "Biomedical Waste Management",
  13: "Health Education and Communication",
  14: "Health Care in India, Health Planning and Management",
  15: "International Health",
  16: "Biostatistics",
  17:"Occupational Health"
};

/*************************************************
 * UTILITIES
 *************************************************/
function normalize(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function expandAbbreviations(text) {
  const map = {
    inh: "tuberculosis",
    rif: "tuberculosis",
    ethambutol: "tuberculosis",
    dots: "tuberculosis",
    imr: "infant",
    tfr: "fertility",
    bmi: "nutrition",
    iec: "healtheducation",
    bcc: "healtheducation",
    lam: "contraception",
    iud: "contraception",
    pm25: "pollution",
    pm10: "pollution",
    tb: "tuberculosis",
    rntcp: "tuberculosis",
    aids:"infection",
    akt:"tuberculosis",
    afp:"flaccid",
    imnci:"programme",
    rch:"programme",
    art:"infection"
  };

  let t = normalize(text);
  Object.keys(map).forEach(k => {
    t = t.replace(new RegExp("\\b" + k + "\\b", "g"), map[k]);
  });
  return t;
}

function isUselessAnswer(a) {
  return /^(a|b|c|d|true|false|none|all of the above|\d+|[ivx]+)$/i
    .test(String(a).trim());
}

/*************************************************
 * OPTION UNIVERSE (READ EVERYTHING EXCEPT Q/A/TITLE)
 *************************************************/
function extractOptionUniverse(row, qCol, aCol, tCol) {
  let parts = [];
  for (let i = 0; i < row.length; i++) {
    if (i === qCol || i === aCol || i === tCol) continue;
    if (row[i]) parts.push(String(row[i]));
  }
  return parts.join(" ");
}

/*************************************************
 * QUESTION PATTERN OVERRIDES (CRITICAL)
 *************************************************/
function questionPatternOverride(question) {
  if (!question) return null;
  const q = normalize(question);

  if (/(cooling power|kata thermometer|wet kata|dry kata)/.test(q)) return 11;
  if (/(wind speed|direction of wind|velocity of wind)/.test(q)) return 11;
  if (/(humidity|relative humidity|ventilation)/.test(q)) return 11;
  if (/(air pollutant|primary air pollutant|secondary air pollutant|not a primary air pollutant)/.test(q)) return 11;

  return null;
}

/*************************************************
 * FREQUENCY BOOST (OPTIONS)
 *************************************************/
function keywordFrequencyBoost(optionsText) {
  if (!optionsText) return null;
  const t = normalize(optionsText);

  if ((t.match(/\bozone\b/g) || []).length >= 2) return 11;
  if ((t.match(/\banemometer\b/g) || []).length >= 2) return 11;
  if ((t.match(/\bhygrometer\b/g) || []).length >= 2) return 11;
  if ((t.match(/\bwet bulb\b/g) || []).length >= 2) return 11;
  if ((t.match(/\bcfcs?\b/g) || []).length >= 2) return 11;

  if ((t.match(/\byellow bag\b/g) || []).length >= 2) return 12;
  if ((t.match(/\bimr\b/g) || []).length >= 2) return 7;

  return null;
}

/*************************************************
 * PHRASE MAP
 *************************************************/
function phraseMap(text) {
  if (!text) return null;
  text = normalize(text);

  if (/(sharp|soiled waste|bmw|colour coding)/.test(text)) return 12;
  if (/(cfcs?|carbon monoxide|greenhouse gases?|humidity|ventilation|air pollution|climate change)/.test(text)) return 11;
  if (/(ethambutol|isoniazid|rifampicin|tuberculosis)/.test(text)) return 5;
  if (/(imr|tfr|replacement level fertility)/.test(text)) return 7;
  if (/(screening|sensitivity|specificity|ppv|npv)/.test(text)) return 4;
  if (/(mean|median|standard deviation|p value)/.test(text)) return 16;
  if (/(health education|counselling|iec|bcc)/.test(text)) return 13;

  return null;
}

/*************************************************
 * KEYWORD ENGINE
 *************************************************/
function detectPSMTopic(text) {
  text = normalize(text);

  if (/(mean|median|mode|standard deviation|variance|p value|confidence interval|chi square|t test|anova|sampling|sample size|type i error|type ii error)/.test(text)) return 16;
  if (/(occupation|occupational|ergonomics|esi|esisc|worker|compensation|bissinosis|anthracosis|bagasosis|silicosis|pneumoconiosis|dust)/.test(text)) return 17;
  if (/(demography|census|birth|contraceptive|pearl|couple|fertility|reproduction|demographic|tfr|imr|mmr|cbr|cdr|population|pyramid|contraception|oral pill|iud|vasectomy|tubectomy|demography|census|birth|contraceptive|contraception|family planning|birth control|pearl|couple|fertility|reproduction|demographic|tfr|imr|mmr|cbr|cdr|population|pyramid|contraception|oral pill|iud|vasectomy|tubectomy|sterilization|tubal ligation|minilap|laparoscopic|intrauterine device|copper t|cuT|combined pill|minipill|progestin only|emergency contraception|medical abortion|mifepristone|misoprostol|injectable|depot|implant|norplant|implanon|jadelle|male condom|female condom|diaphragm|cervical cap|spermicide|foam|gel|natural methods|rhythm method|calendar method|symptothermal method|billings method|ovulation method|cervical mucus method|withdrawal|coitus interruptus|lactational amenorrhea|lam|traditional methods|folk methods|pearl|failure rate|effectiveness|efficacy|continuation rate|couple protection rate|eligible couples|target couples|acceptor|user|unmet need|demand satisfied|contraceptive prevalence|method mix|spacing method|terminal method|limiting method|birth spacing|child spacing|replacement level|demographic transition|population explosion|population pyramid|age structure|sex ratio|dependency ratio|literacy rate|literacy|education|poverty|socioeconomic status|ses|income|wealth|occupation|employment|housing|water|sanitation|electricity|fuel|transport|communication)/.test(text)) return 7;
  if (/(eye|act|national|nhp|blindness|obesity|sputum|registries|vision|acuity|occular|carcinoma|programme|surveillance|nhm|nrhm|nuhm|nvbdcp|ntep|nacp|npcdcs|rmnch|jsy|jssk|icds|ayushman bharat)/.test(text)) return 6;
  if (/(flaccid|tuberculosis|infection|bite|rabies|contamination|fever|vaccine|vaccination|tb|akt|art|imnci|rch|afp|sputum|rntcp|malaria|aids|dengue|hiv|leprosy|diabetes|hypertension|rifampicin|isoniazid|ethambutol|Cancer|cancer|virus|bacteria|bacilus|ethambutol|isoniazid|rifampicin|pyrazinamide|streptomycin|tuberculosis|sputum|smear|culture|dots|rntcp|ntep|short course chemotherapy|directly observed|multidrug resistant|extensively drug resistant|mdr tb|xdr tb|tb vaccine|bcg|mantoux|tuberculin|ghon focus|simon focus|ranke complex|primary complex|miliary tuberculosis|disseminated tuberculosis|extrapulmonary tuberculosis|pleural tuberculosis|lymph node tuberculosis|bone tuberculosis|renal tuberculosis|genitourinary tuberculosis|abdominal tuberculosis|meningeal tuberculosis|tuberculous meningitis|tubercle bacilli|mycobacterium tuberculosis|acid fast bacilli|afb|z n stain|zeihl neelsen stain|auramine phenol stain|chest x ray|ct scan|mri|bronchoscopy|gastric lavage|sputum induction|induced sputum|bronchoalveolar lavage|bal|biopsy|histopathology|cbnaat|cartridge based nucleic acid amplification test|gene xpert|xpert mtb rif|probe|lpa|drug susceptibility testing|dst|culture conversion|treatment completion|cure rate|default rate|failure rate|relapse rate|chemoprophylaxis|regimen|intensive|continuation|category|new case|previously treated|retreatment|failure|default|relapse|chronic|drug resistant|rifampicin resistant|isoniazid resistant|fluoroquinolone resistant|injectable resistant|kanamycin|amikacin|capreomycin|streptomycin|ethionamide|prothionamide|cycloserine|terizidone|pas|para aminosalicylic acid|linezolid|clofazimine|bedaquiline|delamanid|pretomanid|sutezolid|tedizolid|contrepas|levofloxacin|moxifloxacin|gatifloxacin|ofloxacin|ciprofloxacin|delamanid|pretomanid|bedaquiline|linezolid|clofazimine|cycloserine|terizidone|ethionamide|prothionamide|pas|para aminosalicylic acid|streptomycin|kanamycin|amikacin|capreomycin|thioacetazone|thiacetazone|para aminosalicylic acid|cycloserine|terizidone|ethionamide|prothionamide|linezolid|clofazimine|bedaquiline|delamanid|pretomanid|sutezolid|tedizolid|contrepas|levofloxacin|moxifloxacin|gatifloxacin|ofloxacin|ciprofloxacin|delamanid|pretomanid|bedaquiline|linezolid|clofazimine|cycloserine|terizidone|ethionamide|prothionamide|pas|para aminosalicylic acid|streptomycin|kanamycin|amikacin|capreomycin)/.test(text)) return 5;
  if (/(screening|sensitivity|test|reliability|specificity|ppv|npv|lead time bias|wilson jungner)/.test(text)) return 4;
  if (/(incidence|prevalence|attack rate|epidemic|pandemic|cohort|case control|vaccine|immunization|herd immunity)/.test(text)) return 3;
  if (/(energy|pem|xerophthalmia|consumption|dietary|egg|proteins|highest|maiz|rice|anthropometric|meal|bitot|calcium|cow|buffalo|sorghum|pellagra|pellagrogenic|bajara|gram|lathyrism|ergotism|ergot|mid-day|dal|pulses|legumes|zinc|soyabean|beans|metabolic|trace|element|calories|recommended|nutrition|kwashiorkor|marasmus|vitamin|anemia|bmi|stunting|wasting)/.test(text)) return 9;
  if (/(CFCs|noise|hearing|control|scabies|mosquito|aedes|culex|flea|fly|lice|bug|tick|houseflies|water|chlorine|chlorination|sand|filter|mercury|wetbulb|pollutants|heat|stress|Green|acid|rain|velocity|mcardle|sweat|anemometer|hygrometer|temprature|pollution|air|Ozone|greenhouse|gas|noise|pollution|climate|solid waste|humidity|Ventilation|perflation|aspiration|monoxide|greenhouse gases?|humidity|ventilation|air pollution|climate change|noise pollution|heat stress|thermal comfort|mcardle|sweat|anemometer|hygrometer|temperature|pollution|air quality|ozone|carbon|dioxide|nitrogen dioxide|lead|mercury|arsenic|acid rain|smog|photochemical|particulate matter|pm10|pm2|stack height|chimney|plume|activated|sunlight|katathermometer|sling|psychrometer|exhaustion|cramps|hypothermia|frostbite|altitude|decompression|solid|noise|decibel|hearing loss|thermal pollution|radiation|radon|asbestos|mercury|lead|arsenic|fluoride|iodine|iron|water borne|food borne|air borne|vector borne|contact|zoonotic|nosocomial|opportunistic|iatrogenic|idiosyncratic|endemic|epidemic|pandemic|sporadic|outbreak|cluster|exposure|dose|dosage|concentration|threshold|tolerance|susceptibility|resistance|immunity|herd immunity|vaccination|immunization|prophylaxis|chemoprophylaxis|sanitation|hygiene|cleanliness|disinfection|sterilization|fumigation|pest|vector|rodent|insecticide|pesticide|herbicide|fungicide|rodenticide|larvicide|adulticide|space|thermal|fogging|insecticidal|net|llin|indoor|spraying|irs|source|environmental|biological|larvivorous)/.test(text)) return 11;
  if (/(biomedical|recycling|waste|radiation|hazards|yellow|bag|white|blue|incineration|autoclave|sharp|soiled waste|bmw|colour coding|incineration|autoclave|microwave|shredder|biomedical|sharps|container|puncture|leak|tamper|segregation|disposal|landfill|burial)/.test(text)) return 12;
  if (/(antenatal|postnatal|imci|lbw|maternal|preterm|infant|milestone|breast|geriatrics|elderly)/.test(text)) return 8;
  if (/(health education|iec|counselling|motivation|communication)/.test(text)) return 13;
  if (/(who|unicef|world bank|sdg|mdg|international health)/.test(text)) return 15;
  if (/(social|society|accultration|phobia|juvenile|delinquency|moron|retardation|sociology|culture|poverty|literacy|community)/.test(text)) return 10;
  if (/(phc|chc|subcentre|health planning|health economics|budget)/.test(text)) return 14;
  if (/(iceberg|phenomenon|epidemiological|triad|transmitted|transmission|epidemiology|epidemiologist|disease|eradication|eleimination|eradicted|web|levels|prevention|natural|history|disability|index)/.test(text)) return 2;
  if (/(hippocrates|jenner|pasteur|john snow|history of medicine)/.test(text)) return 1;

  return null;
}

/*************************************************
 * MAIN FUNCTION (RUN THIS)
 *************************************************/
function classifyQuestionsPSM_FINAL(sheetName) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName || "Questions");
  if (!sheet) throw new Error("Sheet not found");

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const qCol = headers.findIndex(h => normalize(h) === "question");
  const aCol = headers.findIndex(h => normalize(h).includes("correct"));
  const tCol = headers.findIndex(h => normalize(h).replace(/\s+/g, "") === "formtitle");

  if (qCol === -1 || aCol === -1 || tCol === -1) {
    throw new Error("Required columns missing");
  }

  for (let i = 1; i < data.length; i++) {

    const q = String(data[i][qCol] || "");
    const a = String(data[i][aCol] || "");
    const optionsText = extractOptionUniverse(data[i], qCol, aCol, tCol);

    let topic =
      questionPatternOverride(q) ||
      phraseMap(a) ||
      phraseMap(q) ||
      keywordFrequencyBoost(optionsText) ||
      (!isUselessAnswer(a) && detectPSMTopic(expandAbbreviations(a))) ||
      detectPSMTopic(expandAbbreviations(q + " " + a)) ||
      detectPSMTopic(expandAbbreviations(optionsText));

    sheet.getRange(i + 1, tCol + 1)
      .setValue(topic ? PSM_TOPICS[topic] : "General PSM");
  }

  SpreadsheetApp.getUi().alert("PSM classification completed");
}
