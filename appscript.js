// --- GLOBAL CONSTANTS ---
// By placing these at the top, any function in the script can use them.
const COUNTERPARTIES = [
    "Keystone Dynamics",
    "Weekend Ventures",
    "Iver & Co.",
    "Dancefloor Data",
    "Torches & Co.",
    "Arcade Corp.",
    "Modern Age Ventures",
    "Everlong Enterprises",
    "High Violet Group",
    "Little Lion Holdings",
    "Xyloto Corp.",
    "Dog Days Capital",
    "Electric Feel Energy",
    "Oracular Inc.",
    "Impala Logistics",
    "Coexist Innovations",
    "Innerspeaker Inc.",
    "After Hours Capital",
    "Crystal Castle Holdings",
    "El Camino Group",
    "Wolf Gang Ventures",
    "All Friends Co."
];

// --- Static picklists & Mappings ---
const STANDARDIZED_DOC_TYPE_MAP = {
    "NDA": "Non-Disclosure Agreement (NDA)",
    "MSA": "Master Service Agreement (MSA)",
    "SOW": "Statement Of Work (SOW)",
    "Amendment": "Amendment",
    "Lease": "Lease Agreement",
    "License": "License Agreement",
    "Services": "Services Agreement",
    "Change Order": "Change Order",
    "Purchase": "Purchase Agreement",
    "Consulting": "Consulting Agreement",
    "Offer Letter": "Offer Letter",
    "Employee Separation": "Employee Separation Agreement",
    "Contractor": "Contractor Agreement"
};

const NO_TERM_DOCS = ["Offer Letter", "Employee Separation Agreement"];
const INDUSTRIES = ["Technology", "Healthcare", "Retail", "Manufacturing", "Education"];
const REGIONS = ["NAMER"];
// const REGIONS = ["NAMER", "EMEA", "APAC", "LATAM"];
const TERM_OPTIONS = [1, 2, 3];
const DAYS_NOTICE = [30, 60, 90];
const ASSIGN_OPTS = ["Yes", "No, or consent req’d", "Yes, with conditions"];
const PAYMENT_TERMS = ["30 days", "45 days", "60 days"];

const OBLIGATIONS = {
    "Non-Disclosure Agreement (NDA)": ["Confidentiality", "Data Breach"],
    "Master Service Agreement (MSA)": ["Compliance", "Indemnification", "Insurance", "Limitation of Liability", "Escalation", "DORA"],
    "Statement Of Work (SOW)": ["Deliverables", "Escalation", "Limitation of Liability", "Indemnification", "Insurance"],
    "Lease Agreement": ["Compliance", "Insurance", "Limitation of Liability"],
    "License Agreement": ["Compliance", "Limitation of Liability", "Indemnification"],
    "Services Agreement": ["Service Levels", "Escalation", "Indemnification", "Insurance", "Limitation of Liability"],
    "Change Order": ["Deliverables", "Escalation"],
    "Offer Letter": ["Confidentiality", "Non-Solicitation"],
    "Purchase Agreement": ["Warranties", "Indemnification", "Insurance", "Limitation of Liability"],
    "Consulting Agreement": ["Deliverables", "Indemnification", "Insurance", "Limitation of Liability"],
    "Employee Separation Agreement": ["Confidentiality", "Non-Disparagement"],
    "Contractor Agreement": ["Confidentiality", "Insurance", "Indemnification"]
};

const OBL_TEXT = {
    Compliance: "Compliance: Both parties shall comply with all applicable laws, regulations and internal policies, including anti-corruption, export controls and data privacy rules.",
    Confidentiality: "Confidentiality: All Confidential Information must be protected with at least the same degree of care as the party uses for its own, and not less than reasonable care.",
    "Data Breach": "Data Breach: Each party will notify the other within 48 hours of any security incident affecting personal data and cooperate on remediation efforts.",
    Deliverables: "Deliverables: All deliverables must meet the acceptance criteria defined in the SOW and be delivered in both draft and final formats.",
    Escalation: "Escalation: Unresolved issues will escalate to executive sponsors within 5 business days, following the path: Project Manager → VP → CEO.",
    Indemnification: "Indemnification: Each party will indemnify the other against third-party claims for breach, negligence or IP infringement, subject to notice and defense requirements.",
    Insurance: "Insurance: Contractor shall maintain general liability ($1M per occurrence), professional liability ($2M aggregate) and workers’ comp coverage as required by law.",
    "Limitation of Liability": "Limitation of Liability: Neither party’s aggregate liability will exceed the total fees paid under this agreement, except for willful misconduct or gross negligence.",
    "Service Levels": "Service Levels: Provider guarantees 99.9% uptime and will credit fees for any monthly downtime exceeding SLA targets.",
    Warranties: "Warranties: Seller warrants that goods will conform to specifications for 12 months and will repair or replace defective items at no cost.",
    "Non-Solicitation": "Non-Solicitation: For 12 months post-termination, neither party will solicit or hire the other’s employees or contractors.",
    "Non-Disparagement": "Non-Disparagement: Both parties agree not to make adverse or negative public statements about the other following separation.",
    DORA: "DORA: Parties shall comply with the EU Digital Operational Resilience Act..."
};

/**
 * Processes a sample document request from a spreadsheet row.
 * Reads inputs, determines the workflow (sets or random), generates documents,
 * saves them to Google Drive, and provides feedback in a "Status" column.
 *
 * @param {Object} e The event object from the onEdit trigger.
 */
function submitSampleRequest(e) {
    // Check if the trigger is a valid sheet edit
    if (!e || !e.source) {
        Logger.log("Function called without a valid event trigger.");
        return;
    }
    const pick = arr => arr[Math.floor(Math.random() * arr.length)];

    const spreadsheet = e.source;
    const sheet = spreadsheet.getActiveSheet();
    const activeRange = spreadsheet.getActiveRange();
    const firstRow = activeRange.getRow();

    // Exit if the edit is on the wrong sheet or in the header row
    if (sheet.getName() !== "Requests" || firstRow <= 1) {
        return;
    }

    // --- CONFIGURATION ---
    const COLUMN_MAP = {
        EMAIL: 2,
        QUANTITY: 3,
        SPECIAL_INSTRUCTIONS: 4,
        DOC_TYPES: 5,
        LANGUAGE: 6,
        FIRST_PARTY: 7,
        CREATE_SETS: 8,
        STATUS: 9
    };
    const statusRange = sheet.getRange(firstRow, COLUMN_MAP.STATUS);

    try {
        statusRange.setValue("Processing...");

        // Securely get Root Folder ID from Script Properties
        const properties = PropertiesService.getScriptProperties();
        const rootFolderId = properties.getProperty('ROOT_FOLDER_ID');
        if (!rootFolderId) {
            throw new Error("ROOT_FOLDER_ID is not set in Script Properties.");
        }

        // Read all data from the spreadsheet row into an object
        const requestData = {
            email: sheet.getRange(firstRow, COLUMN_MAP.EMAIL).getValue(),
            quantity: parseInt(sheet.getRange(firstRow, COLUMN_MAP.QUANTITY).getValue(), 10),
            specialInstructions: sheet.getRange(firstRow, COLUMN_MAP.SPECIAL_INSTRUCTIONS).getValue(),
            docTypeString: sheet.getRange(firstRow, COLUMN_MAP.DOC_TYPES).getValue(),
            language: sheet.getRange(firstRow, COLUMN_MAP.LANGUAGE).getValue(),
            firstParty: sheet.getRange(firstRow, COLUMN_MAP.FIRST_PARTY).getValue() || "Elston Enterprises",
            createSets: sheet.getRange(firstRow, COLUMN_MAP.CREATE_SETS).getValue() === true
        };

        if (!requestData.email || !requestData.quantity || requestData.quantity <= 0) {
            throw new Error("Invalid or missing Email or Quantity.");
        }

        // --- FOLDER SETUP ---
        const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
        const folderName = `${requestData.email}_${timestamp}`;
        const rootFolder = DriveApp.getFolderById(rootFolderId);
        const subfolder = rootFolder.createFolder(folderName);

        // --- CONDITIONAL WORKFLOW ---
        let successMessage = "";

        if (requestData.createSets === true) {
            // WORKFLOW 1: CREATE DOCUMENT SETS
            const numSets = Math.floor(requestData.quantity / 5);
            const docCount = numSets * 5;
            const DOCUMENT_SET_TYPES = ["Non-Disclosure Agreement (NDA)", "Master Service Agreement (MSA)", "Statement Of Work (SOW)", "Statement Of Work (SOW)", "Change Order"];

            if (numSets < 1) {
                throw new Error("Quantity must be 5 or more to create sets.");
            }

            for (let i = 0; i < numSets; i++) {
                // CORRECT: `parentMsaData` is declared here, visible to the entire set.
                let parentMsaData = {};
                const setCounterparty = pick(COUNTERPARTIES);
                for (const docType of DOCUMENT_SET_TYPES) {
                    const docData = generateSetDocumentRow(requestData, docType, setCounterparty);

                    if (docType.includes("MSA")) {
                        docData.contractNumber = generateContractNumber();
                        // This assigns the MSA data to the variable.
                        parentMsaData = docData;
                    } else if (docType.includes("SOW") || docType.includes("Change Order")) {
                        // This will now work, because `parentMsaData` was defined outside this block
                        // and has already been assigned the MSA's data.
                        docData.parentMsaContractNumber = parentMsaData.contractNumber;
                        docData.parentMsaDate = Utilities.formatDate(parentMsaData.effectiveDate, Session.getScriptTimeZone(), "MM/dd/yyyy");
                    }
                    Logger.log("Doc Data: " + docData);
                    processAndCreateFile(docData, subfolder);
                }
            }
            successMessage = `Success! ${docCount} documents (${numSets} sets) created.`;

        } else {
            // WORKFLOW 2: CREATE INDIVIDUAL RANDOM DOCUMENTS
            const docCount = requestData.quantity;
            for (let i = 0; i < docCount; i++) {
                const docData = generateRandomDocumentRow(requestData);
                processAndCreateFile(docData, subfolder);
            }
            successMessage = `Success! ${docCount} individual documents created.`;
        }

        // --- SUCCESS ---
        statusRange.setValue(successMessage);
        sendSlackNotification(requestData.email, successMessage, requestData.language, subfolder.getUrl());
        Logger.log("Slack notification sent with folder link.");

    } catch (error) {
        // Catch any error and write it to the Status column
        Logger.log(`Error processing request in row ${firstRow}: ${error.message}`);
        statusRange.setValue(`Error: ${error.message}`);
    }
}

/**
 * Builds the core array of details (dates, terms, etc.) for an agreement.
 * @param {string} agreementType The type of agreement being generated.
 * @return {Object} An object containing the effectiveDate and an array of detail strings.
 */
function buildAgreementDetails(agreementType) {
    const today = new Date();
    const pick = arr => arr[Math.floor(Math.random() * arr.length)];
    const addDays = (date, days) => new Date(date.getTime() + days * 24 * 60 * 60 * 1000);
    const fmt = date => Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy");

    const isNoTerm = NO_TERM_DOCS.includes(agreementType);
    let effectiveDate, termYears, termEndDate, renewalNoticePeriod, renewalNoticeDate, actionRequiredBy;

    if (isNoTerm) {
        const daysAgo = pick([30, 60, 90, 180, 365]);
        effectiveDate = addDays(today, -daysAgo);
    } else {
        const daysUntilExpiry = Math.floor(Math.random() * 180) + 1;
        termEndDate = addDays(today, daysUntilExpiry);
        termYears = pick(TERM_OPTIONS);
        effectiveDate = new Date(termEndDate);
        effectiveDate.setFullYear(effectiveDate.getFullYear() - termYears);
        renewalNoticePeriod = pick(DAYS_NOTICE);
        renewalNoticeDate = addDays(termEndDate, -renewalNoticePeriod);
        actionRequiredBy = addDays(today, Math.floor(Math.random() * 180));
    }

    const parts = [`Effective Date: ${fmt(effectiveDate)}`];

    if (!isNoTerm) {
        parts.push(
            `Initial Term: ${termYears} year(s)`,
            `Expiration Date: ${fmt(termEndDate)}`,
            `Renewal Notice Period: ${renewalNoticePeriod} days`,
            `Renewal Notice Date: ${fmt(renewalNoticeDate)}`,
            `Action Required By: ${fmt(actionRequiredBy)}`,
            `Assignment (General): ${pick(ASSIGN_OPTS)}`,
            `Assignment (Change of Control): ${pick(ASSIGN_OPTS)}`,
            `Assignment (Termination Rights): ${pick(["Yes", "No"])}`,
            `Payment Terms: ${pick(PAYMENT_TERMS)}`,
            `Termination for Cause Notice: ${pick(DAYS_NOTICE)} days`,
            `Termination for Convenience Notice: ${pick(DAYS_NOTICE)} days`
        );
    }

    // Return an object containing both the date object and the parts array
    return { effectiveDate: effectiveDate, parts: parts };
}

/**
 * Generates a data OBJECT for a specific document within a set.
 * @param {Object} requestData An object containing the user's request details.
 * @param {string} agreementType The specific type of agreement to generate.
 * @param {string} counterparty The specific counterparty for this document set.
 * @return {Object} An object containing all generated data for the document.
 */
function generateSetDocumentRow(requestData, agreementType, counterparty) {
    // Helpers
    const pick = arr => arr[Math.floor(Math.random() * arr.length)];
    const shuffle = arr => arr.sort(() => Math.random() - 0.5);
    const addDays = (date, days) => new Date(date.getTime() + days * 24 * 60 * 60 * 1000);
    const fmt = date => Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy");

    // --- Get Core Details ---
    // Call the helper, which now returns an object with the date and parts array
    const detailsObject = buildAgreementDetails(agreementType);
    const parts = detailsObject.parts;

    // --- Add SOW-specific details if applicable ---
    if (agreementType.includes("SOW")) {
        const today = new Date();
        const totalValue = Math.floor(Math.random() * 450000) + 50000;
        const depositAmount = Math.floor(Math.random() * 20000) + 5000;
        const oneTimeAmount = Math.floor(Math.random() * 40000) + 10000;
        const depositDue = addDays(today, Math.floor(Math.random() * 180));
        const oneTimeDue = addDays(today, Math.floor(Math.random() * 180));

        parts.push(
            `Total Contract Value: $${totalValue.toLocaleString()} USD`,
            `Deposit Amount: $${depositAmount.toLocaleString()} USD, Deposit Due: ${fmt(depositDue)}`,
            `One-Time Payment: $${oneTimeAmount.toLocaleString()} USD, Due: ${fmt(oneTimeDue)}`
        );
    }

    // --- Add Obligation Text ---
    const possible = OBLIGATIONS[agreementType] || []; // Uses global constant
    const selected = shuffle(possible).slice(0, pick([1, 2, 3]));
    selected.forEach(key => {
        if (OBL_TEXT[key]) { // Uses global constant
            parts.push(OBL_TEXT[key]);
        }
    });

    // --- Assemble and Return Final Data Object ---
    const docData = {
        email: requestData.email,
        language: requestData.language,
        firstParty: requestData.firstParty,
        counterparty: counterparty,
        agreementType: agreementType,
        industry: pick(INDUSTRIES),   // Uses global constant
        geography: pick(REGIONS),     // Uses global constant
        specialInstructions: parts.join(", "),
        effectiveDate: detailsObject.effectiveDate // The raw date object
    };

    return docData;
}

/**
 * Generates data for a single, random document using global constants.
 * @param {Object} requestData An object containing the user's request details.
 * @return {Object | null} An object containing all generated data for the document, or null.
 */
function generateRandomDocumentRow(requestData) {
    // Helpers
    const pick = arr => arr[Math.floor(Math.random() * arr.length)];
    const shuffle = arr => arr.sort(() => Math.random() - 0.5);
    const addDays = (date, days) => new Date(date.getTime() + days * 24 * 60 * 60 * 1000);
    const fmt = date => Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy");

    // --- Select Random Document Type & Parties ---
    const firstParty = requestData.firstParty;
    const counterparty = pick(COUNTERPARTIES); // Uses global constant

    Logger.log(`Generating random document for ${firstParty} with counterparty ${counterparty}`);

    const extractedTypes = extractDocTypes(requestData.docTypeString);
    const standardizedTypes = extractedTypes
        .map(t => STANDARDIZED_DOC_TYPE_MAP[t] || t) // Uses global constant
        .filter(t => !!OBLIGATIONS[t]); // Uses global constant

    if (standardizedTypes.length === 0) {
        Logger.log("No valid document types found for random generation in: " + requestData.docTypeString);
        return null;
    }
    const agreementType = pick(standardizedTypes);

    // --- Build Core Details ---
    const detailsObject = buildAgreementDetails(agreementType);
    const parts = detailsObject.parts;

    // --- Add SOW-specific details if applicable ---
    if (agreementType.includes("SOW")) {
        const today = new Date();
        const totalValue = Math.floor(Math.random() * 450000) + 50000;
        const depositAmount = Math.floor(Math.random() * 20000) + 5000;
        const oneTimeAmount = Math.floor(Math.random() * 40000) + 10000;
        const depositDue = addDays(today, Math.floor(Math.random() * 180));
        const oneTimeDue = addDays(today, Math.floor(Math.random() * 180));

        parts.push(
            `Total Contract Value: $${totalValue.toLocaleString()} USD`,
            `Deposit Amount: $${depositAmount.toLocaleString()} USD, Deposit Due: ${fmt(depositDue)}`,
            `One-Time Payment: $${oneTimeAmount.toLocaleString()} USD, Due: ${fmt(oneTimeDue)}`
        );
    }

    // --- Add Obligation Text ---
    const possible = OBLIGATIONS[agreementType] || [];
    const selected = shuffle(possible).slice(0, pick([1, 2, 3]));
    selected.forEach(key => {
        if (OBL_TEXT[key]) {
            parts.push(OBL_TEXT[key]);
        }
    });

    // --- Assemble and Return Final Data Object ---
    const docData = {
        email: requestData.email,
        language: requestData.language,
        firstParty: firstParty,
        counterparty: counterparty,
        agreementType: agreementType,
        industry: pick(INDUSTRIES),
        geography: pick(REGIONS),
        specialInstructions: parts.join(", "),
        effectiveDate: detailsObject.effectiveDate
    };

    return docData;
}

/**
 * Assembles the complete AI prompt by calling all modular helper functions.
 * @param {Object} docData An object containing all data for a single document.
 * @return {string} The final, complete prompt to be sent to the AI.
 */
function createPrompt(docData) {
    // Destructure all needed properties from the docData object
    const {
        agreementType, industry, geography, language, specialInstructions,
        firstParty, counterparty, parentMsaContractNumber, parentMsaDate, contractNumber
    } = docData;

    Logger.log(`Inside createPrompt - agreementType=${agreementType}, firstParty=${firstParty}, parentMsaNum=${parentMsaContractNumber}, contractNumber=${contractNumber}`);
    // --- Conditionally create a reference to the parent MSA ---
    let msaReference = "";
    if (parentMsaContractNumber) {
        msaReference = `This document is governed by the Master Service Agreement (MSA) between the parties, dated ${parentMsaDate}, with Contract Number ${parentMsaContractNumber}. Ensure the terms of this document are consistent with the parent MSA.`;
    }

    // --- Conditionally create a display for the doc's own contract number ---
    let contractNumberDisplay = "";
    if (contractNumber) {
        contractNumberDisplay = `Contract Number: ${contractNumber}`;
    }

    // --- Call all the modular helper functions ---
    const setup = getPromptSetup(agreementType, industry, geography);
    const objective = getPromptObjective(industry, geography);
    const instructions = getPromptInstructions();
    const generalRequirements = getPromptGeneralRequirements();
    const outputFormat = getPromptOutputFormat();
    const specificRequirements = getPromptSpecifics(agreementType, firstParty, counterparty);

    // --- Assemble the final prompt ---
    const documentSpecificRequirements = `Document-Specific Requirements:
${specificRequirements}
For each document type, ensure that all sections are written as complete paragraphs, with no bullet points or numbered lists. Each section should flow naturally, fully explaining the legal concepts and providing clarity to the terms. Each document should reflect ${firstParty} business and legal requirements as outlined. 
`;

    const prompt = `
${msaReference}
${contractNumberDisplay}

Inputs:
Agreement Type: [${agreementType}]
Industry: [${industry}]
Geography: [${geography}]
Language: [${language}]
Special Instructions: [${specialInstructions}]

${setup}
${objective}
${instructions}
${generalRequirements}
${outputFormat}

${documentSpecificRequirements}

Instructions:

1. Research agreements for the input Industry and Geography
2. Review the Special Instructions. 
3. Use this information to generate an agreement that provides a realistic representation of this type of agreement. DO NOT include any explanatory text before or after the agreement. 
`;

    console.log("Doc Specific Reqs: " + documentSpecificRequirements);
    return prompt;
}

function getPromptSetup(agreementType, industry, geography) {
    const setup = `What it means to be a Sample Document Assistant
This Sample Document Assistant is an expert in creating tailored legal documents containing realistic sample data and language. The assistant can generate agreements for any industry or geography, and will research as necessary to create sample agreements for demonstrations. The goal of these demonstrations is to ensure the audience that Docusign understands their business, and a realistic document including realistic language can help achieve that goal.

Output
Generate a .docx file that serves as a sample agreement for the specified type of legal agreement. This agreement will be tailored for a ${agreementType} for a ${industry} business in ${geography}, will be written in the language provided in the input, and include sample information for fictitious agreements. `;
    return setup;
}

function getPromptInstructions(firstParty) {
    return `Instructions for Using the Variables:
[Industry] Variable:
The [industry] variable represents the type of business or sector in which ${firstParty} operates. When using the script, [industry] should be replaced with the appropriate business type relevant to the context of the contract.
For example:
If ${firstParty} operates in software development, replace [industry] with "software development".
If ${firstParty} is in solar panel installation and services, replace [industry] with "solar panel services".
[Country] Variable:
 The [country] variable should be replaced with only one country name where ${firstParty} is based or where the agreement is being applied. This will ensure that the document’s content, including regulations, currency, and date format, aligns with the legal requirements of the country.
 For example:
If the company is based in France, replace [country] with "France".
If the agreement is being generated for a UK-based contract, replace [country] with "United Kingdom".
For any other countries, replace [country] accordingly.
Fictitious Company Address:
 The Fictitious Company Address will be automatically generated based on the [country] variable. For example:
For France, the address might be "123 Rue de la Technologie, Paris, 75001".
For United Kingdom, the address might be "45 High Street, London, EC1A 1BB".
 The generated company address will ensure a credible and localized appearance for the fictitious company.
Execution Date:
 The Execution Date will be automatically generated within the current year, in the format of dd/mm/yyyy. The date should range from 01/01/yyyy to 31/12/yyyy (e.g., 15/11/2024).
Jurisdiction City:
 The Jurisdiction City for all documents will be the capital of the country specified in the [country] variable. For example:
France: Paris
United Kingdom: London
Germany: Berlin
United States: Washington, D.C.
  `;
}

function getPromptObjective(firstParty) {

    const objective = `Objective:
To generate several Word documents with contractual information in [Language]. These documents should reflect the specific context of ${firstParty}, a company based in [Country], operating in the [industry] sector. Each document must be comprehensive, realistic, compliant with professional legal practices, and adapted to the company’s activities in [industry]. Where applicable, the documents must include relevant local regulations specific to the [industry] and [country], such as BaFin in Germany or GDPR in the EU. Additionally, all monetary values in the documents must use the currency consistent with the country where the agreement applies (e.g., euros for France, pounds for the UK, etc.). Any numbers should be written out in words, followed by the numerical form in parentheses, wherever possible (e.g., "thirty (30) days").Dates should be formatted according to the [country]’s standard date format (e.g., dd/mm/yyyy for France, mm/dd/yyyy for the United States, etc.).

  `;
    return objective;
}

function getPromptGeneralRequirements(firstParty, counterparty) {
    return `General Requirements:

### Agreement Structure
The generated agreement must follow a logical and professional structure. Adhere to the following architectural guidelines:

1.  **Preamble and Recitals**: Begin with a preamble identifying the parties (${firstParty}, ${counterparty}) and the Effective Date. Follow this with "Whereas" clauses (Recitals) that explain the business context and the purpose of the agreement.

2.  **Definitions Section**: Immediately after the Recitals, include a numbered "Definitions" section. All key terms that are used in multiple sections of the agreement (e.g., "Confidential Information," "Services," "Term") must be defined here.

3.  **Logical Section Flow**: The body of the agreement should follow this conventional order:
    * Core Business Terms (e.g., Scope of Services, Payment, Deliverables).
    * Term and Termination provisions.
    * Representations and Warranties.
    * Covenants and Ongoing Obligations (e.g., Confidentiality).
    * Risk Allocation (e.g., Indemnification, Limitation of Liability).
    * Miscellaneous / Boilerplate Provisions.
    * Signatures and Exhibits.

4.  **Boilerplate Section**: Group all standard legal clauses (e.g., Governing Law, Notices, Assignment, Force Majeure, Entire Agreement, Severability) into a final section titled "Miscellaneous" or "General Provisions".

5.  **Signatures**: Conclude with a proper signature block for the authorized representatives of both ${firstParty} and ${counterparty}, including lines for name, title, and date.

### Content and Style Requirements
Level of Detail:
 All content must be extensively detailed and realistic. Each section must include several lines of paragraphs written in formal legal language. No bullet points or numbered lists are allowed in the content. Information should be presented in long paragraphs, fully explaining all terms and concepts.

Contextualization: Every section must thoroughly explain the terms and clauses, illustrated with real-world examples or industry-specific use cases where applicable.

Key Terms: All terms should be finalized, avoiding placeholders or incomplete information.

Industry-Specific Adaptation:  Each agreement should reference practices, regulations, and compliance requirements specific to [industry], making the documents contextually relevant to ${firstParty}’s operations.

Local Regulations: Relevant local laws, regulations, or authorities (e.g., BaFin for financial services in Germany, GDPR for data protection in the EU, or FCA in the UK) must be included where appropriate for the [industry] and [country].

Currency and Date Formatting: All monetary values and dates must use the appropriate format for the specified [country].

Writing Style:
The language must be professional, formal, and employ appropriate legal terminology. Use conditional and subordinate clauses to ensure all provisions are comprehensive, protective, and unambiguous. Include binding language affirming the agreement applies to both parties and their successors.
`;
}

function getPromptSpecifics(agreementType, firstParty, counterparty) {

    let specificRequirements = '';

    switch (agreementType) {
        case 'Non-Disclosure Agreement (NDA)':
            specificRequirements = `1. Non-Disclosure Agreement (NDA) Structure :
Introduction: The NDA must explain the purpose of the agreement, which is to protect confidential information exchanged between the parties, particularly relating to [industry] services (e.g., software solutions, solar panel services) provided by ${firstParty}.
Parties Involved: The agreement is between ${firstParty} and ${counterparty}, with full legal details such as registration number and address.
Definition of Confidential Information: A detailed explanation of what constitutes confidential information. This should include proprietary technical data, software code, algorithms, business strategies, and client information shared in confidence by ${firstParty}.
Obligations of the Parties: Extensive obligations should be outlined to ensure confidentiality and to limit the use of shared information solely for the purpose of evaluating or carrying out the business relationship.
Duration of Confidentiality Obligation: Typically, five (5) years after termination of the agreement, unless a different period is agreed upon.
Exceptions to Confidentiality: Publicly available information, information that becomes available to the receiving party through lawful means from a third party, or information that is independently developed without reference to the confidential data.
Data Protection (GDPR): Compliance with GDPR must be explicitly included.
Penalties for Breach: Financial and legal consequences for violating confidentiality obligations.
Applicable Law and Jurisdiction: Jurisdiction under [Country] law, with dispute resolution mechanisms.
Signatures: Space for names, titles, and signatures of authorized representatives for both parties.`;
            break;
        case 'Master Service Agreement (MSA)':
            specificRequirements = `2. Master Service Agreement (MSA) Structure :
Introduction: Explains the roles of ${firstParty} and ${counterparty}, outlining the scope of services provided (e.g., software support, solar panel maintenance).
Scope of Services: Detailed descriptions of services provided, including maintenance, consulting, and delivery timelines.
Orders and Statements of Work (SOW): Details on how Statements of Work (SOW) will be issued, outlining deliverables, costs, and performance standards.
Payment Terms: Payment schedules, penalties for delays, billing terms, and currency specifications. Include terms such as Annual Contract Value, Total Contract Value, Late Fees Apply, and Late Fee % where applicable.
Assignment and Change of Control:
  • Assignment (General): Conditions under which the agreement may be assigned or transferred.  
  • Assignment (Change of Control): Special provisions related to assignment during a merger or acquisition.  
  • Assignment (Termination Rights): Conditions under which assignment rights may change following contract termination.
Intellectual Property: Ownership clauses for software and custom developments.
Data Protection and Cybersecurity: Compliance with GDPR and specific cybersecurity measures.
Force Majeure: Detailed scenarios for force majeure events, including steps for resolution.
Dispute Resolution: Multi-step processes including mediation and arbitration.
Limitation of Liability: Include terms such as Liability Cap Amount, Liability Cap Duration, and Liability Cap Multiplier to specify the limits of financial responsibility in case of breach.
Governing Law and Jurisdiction: Detailed information on Governing Law, Jurisdiction, and dispute resolution procedures, with jurisdiction based on the [country] capital city.
ESG Clause (Environmental, Social, and Governance): Provisions for environmental responsibility, social impact, and governance practices (e.g., sustainability targets, diversity and inclusion, environmental compliance for solar panels).
Signatures: Space for authorized representatives of both parties to sign.
DORA Compliance: Establish and maintain an ICT risk-management framework aligned with DORA, including semi-annual policy reviews and resilience testing. Require notification timelines and root-cause analysis obligations per DORA Articles 22-23. Mandate annual digital resilience exercises (e.g., penetration tests, threat simulations) with summary reporting requirements. Ensure all DORA-related requirements survive non-renewal, termination, or liability provisions.
`;
            break;
        case 'Statement Of Work (SOW)':
            specificRequirements = `3. Statement Of Work (SOW) Structure :
Introduction: Reference the underlying Master Agreement and explain the purpose of this SOW—to define deliverables and responsibilities for ${firstParty} (first party) and ${counterparty} (counter party).
Project Overview: High-level summary of objectives, goals and business drivers.
Parties Involved: Full legal names, addresses and contact info for both Client and Contractor.
Scope of Work & Deliverables: Detailed list of tasks, milestones, deliverables and acceptance criteria.
Timeline & Milestones: Start/end dates, phase gates, review meetings, and final acceptance date.
Change Control Process: How scope changes are proposed, approved, and priced (reference Change Order process).
Fees & Payment Schedule: Rates, payment triggers (e.g., upon milestone acceptance), invoicing frequency, and expense reimbursement.
Client Responsibilities: Data, resources or approvals the Client must provide to enable timely delivery.
Confidentiality & Data Security: Reference to NDA or specific data-protection obligations.
Governing Law & Dispute Resolution: Choice of law, venue, mediation/arbitration steps.
Signatures: Signature blocks for authorized representatives of both parties.`;
            break;
        case 'Amendment':
            specificRequirements = `4. Amendment Structure :
            Parties Involved: The agreement is between ${firstParty} and ${counterparty}, with full legal details such as registration number and address.
Original Agreement Reference: Identify the original agreement (name, date, parties) being amended.
Recitals: Brief statements explaining the background and purpose of the amendment.
Amending Provisions: Exactly which clauses or sections are changed—quote old text and provide new text.
Effective Date: The date on which the amendment takes effect.
No Other Changes: A clause confirming that all other terms of the original agreement remain in full force.
Integration Clause: Statement that this amendment and the original document together form the complete agreement.
Governing Law & Jurisdiction: Confirm that the same choice of law/venue applies.
Signatures & Dates: Signature lines and dates for both parties’ authorized representatives.`;
            break;
        case 'Lease':
            specificRequirements = `5. Lease Agreement Structure :
            Parties Involved: The agreement is between ${firstParty} and ${counterparty}, with full legal details such as registration number and address.
Premises Description: Exact address and square footage of the leased space.
Term & Renewal Options: Lease commencement, expiration, any renewal terms and notice periods.
Rent & Additional Charges: Base rent, escalation clauses, CAM charges, utilities, taxes, and security deposit amount.
Use of Premises: Permitted uses, prohibited activities, compliance with zoning and laws.
Maintenance & Repairs: Landlord vs. tenant responsibilities for repairs and upkeep.
Alterations & Improvements: Process for approval of tenant improvements, restoration obligations at lease end.
Insurance Requirements: Types and limits of insurance each party must carry.
Default & Remedies: Events of default, cure periods, and landlord’s remedies.
Assignment & Subletting: Conditions under which tenant may assign or sublet.
Termination: Early termination rights, penalties, and holdover provisions.
Governing Law & Venue: Applicable law and dispute-resolution mechanism.
Signatures: Authorized signature blocks for landlord and tenant.`;
            break;
        case 'License':
            specificRequirements = `6. License Agreement Structure :
            Parties Involved: The agreement is between ${firstParty} and ${counterparty}, with full legal details such as registration number and address.
Grant of License: Define scope (rights granted, territory, exclusivity, duration).
Restrictions: Prohibited actions (e.g., no sublicensing, reverse engineering).
Consideration & Payment: License fees, payment schedule, audit rights.
Intellectual Property Ownership: Affirm that licensor retains all IP rights.
Support & Updates (if applicable): Describe any maintenance or upgrade commitments.
Warranties & Disclaimers: Licensor’s warranties and buyer’s disclaimer of implied warranties.
Indemnification: Who indemnifies whom for third-party claims.
Limitation of Liability: Cap on damages and excluded liabilities.
Termination: Events causing termination, post-termination rights/responsibilities.
Confidentiality & Data Security: Any data-handling rules or NDA cross-reference.
Governing Law & Venue: Applicable jurisdiction.
Signatures: Authorized representatives of licensor and licensee.`;
            break;
        case 'Services Agreement':
            specificRequirements = `7. Services Agreement Structure :
            Parties Involved: The agreement is between ${firstParty} and ${counterparty}, with full legal details such as registration number and address.
Introduction: Identify the parties and the overall intent—to deliver [type of services].
Scope of Services: Detailed description of services, deliverables, and service levels/KPIs.
Duration & Termination: Contract period, renewal terms, termination for convenience and for cause.
Fees & Expenses: Pricing model (fixed, time & materials), invoicing schedule, expense reimbursement.
Change Management: How additional work is requested, approved, and billed.
Client Obligations: Access, resources, or cooperation required from the client.
Confidentiality & Data Protection: NDA cross-reference and GDPR/CCPA compliance.
Intellectual Property: Ownership of work product and any background IP licenses.
Warranties & Representations: Service quality commitments and disclaimers.
Indemnification: Protection against third-party claims.
Limitation of Liability: Caps and exclusions.
Insurance: Required coverages and limits.
Dispute Resolution: Mediation/arbitration process.
Governing Law & Venue: Choice of law.
Signatures: Blocks for both parties.`;
            break;
        case 'Change Order':
            specificRequirements = `8. Change Order Structure :
            
Contract Reference: Cite the original SOW or Services Agreement by name and date.
Parties Involved: The agreement is between ${firstParty} and ${counterparty}, with full legal details such as registration number and address.
Change Description: Detailed narrative of scope changes, new tasks or omissions.
Impact Analysis: Effects on timeline, milestones and deliverables.
Revised Schedule & Milestones: Updated dates for each deliverable or phase.
Pricing Adjustments: Additional or reduced fees, payment schedule changes.
Approval Process: Sign-off lines, approval workflow, and effective date.
Integration with Original Agreement: Statement that all other terms remain unchanged.
Governing Law & Venue: Same jurisdiction as the original contract.
Signatures & Dates: Authorized signatures for both parties.`;
            break;
        case 'Offer Letter':
            specificRequirements = `9. Offer Letter Structure :
            Parties Involved: The agreement is between ${firstParty} and an individual prospective employee.
Position & Reporting: Job title, department, manager name and work location.
Start Date & Employment Type: Full-time/part-time, exempt/non-exempt, at-will statement.
Compensation: Base salary, bonus eligibility, equity grant details (if any).
Benefits & Perks: Health, retirement, PTO, relocation assistance, etc.
Conditions of Employment: Background check, drug test, verification of eligibility.
Confidentiality & IP Assignment: Standard clauses for proprietary information and inventions.
Probationary Period (if any): Duration and evaluation process.
Acceptance Instructions: How and by when to sign and return.
Governing Law: Employment law jurisdiction.
Signature Block: Candidate and hiring manager signatures and dates.
Include sign-on bonus with a clawback/repayment option based upon term of employment (i.e. you have to work for 6/12/24 months or your sign-on bonus is clawed back.)`;
            break;
        case 'Purchase Agreement':
            specificRequirements = `10. Purchase Agreement Structure :
            Parties Involved: The agreement is between ${firstParty} and ${counterparty}, with full legal details such as registration number and address.
Parties & Definitions: Buyer, Seller and any key terms (e.g., “Products,” “Services”).
Sale & Purchase: Detailed description of goods/services, quantities, quality standards.
Purchase Price & Payment Terms: Price, currency, taxes, payment schedule, late fees.
Delivery & Acceptance: Incoterms (if applicable), delivery location, inspection and rejection rights.
Title & Risk of Loss: When title and risk pass from seller to buyer.
Warranties & Remedies: Seller warranties, remedy process, cure periods.
Taxes & Duties: Allocation of sales/use tax, customs duties.
Confidentiality (if applicable): Protection of pricing and proprietary specs.
Indemnification: Protection against third-party claims.
Limitation of Liability: Damage caps and excluded liabilities.
Termination & Suspension: Conditions for termination, effect on orders.
Governing Law & Dispute Resolution: Venue and governing law.
Signatures: Authorized signatures and dates for both parties.`;
            break;
        case 'Consulting Agreement':
            specificRequirements = `11. Consulting Agreement Structure :
            Parties Involved: The agreement is between ${firstParty} and ${counterparty}, with full legal details such as registration number and address.
Engagement Scope: Define services to be provided by Consultant, deliverables, and objectives.
Term & Termination: Start date, duration, renewal options, and termination rights.
Fees & Payment: Consulting rates (hourly/project), invoicing schedule, payment terms, and expense reimbursement.
Client Obligations: Data access, approvals, resources or personnel support required.
Confidentiality & Non-Disclosure: Protection of client information and IP.
Work Product Ownership: Assignment of IP created during engagement.
Indemnification & Insurance: Liability protection and required coverage limits.
Independent Contractor Status: Statement that Consultant is not an employee and responsible for taxes.
Conflict of Interest: Duty to disclose any conflicts.
Warranties & Representations: Service quality commitments.
Limitation of Liability: Caps on damages.
Non-Solicitation: Restrictions on poaching client’s employees.
Governing Law & Venue: Applicable jurisdiction.
Signatures: Authorized blocks for both parties.`;
            break;
        case 'Employee Separation Agreement':
            specificRequirements = `12. Employee Separation Agreement Structure :
Parties & Effective Date: ${firstParty} and Employee names, and separation effective date.
Separation Terms: Last day of employment, pay in lieu of notice, severance package details.
Release of Claims: Employee’s waiver of legal claims in exchange for consideration.
Confidentiality & Non-Disclosure: Ongoing obligations to protect company information.
Non-Disparagement: Mutual promises not to make negative statements.
Return of Property: Process for returning company assets (laptop, keys, etc.).
Benefits & COBRA: Continuation of benefits, COBRA rights, and timing.
Reference & Rehire Rights: Any agreed reference language or rehire eligibility.
Tax Treatment: How severance is taxed and any gross-up provisions.
Governing Law & Dispute Resolution: Venue and applicable law.
Signatures: Employee and authorized company representative.`;
            break;
        case 'Contractor Agreement':
            specificRequirements = `13. Contractor Agreement Structure :
Engagement & Scope: Services to be performed by Contractor, deliverables and milestones.
Term & Renewal: Agreement period, renewal options, and termination for convenience or cause.
Compensation & Expenses: Fees (fixed, hourly or milestone-based), invoicing schedule, and expense reimbursement.
Independent Contractor Status: Explicit statement of non-employee relationship and tax responsibilities.
Confidentiality & IP Assignment: Protection of sensitive info and assignment of work-product IP.
Compliance with Laws: Contractor’s duty to comply with applicable regulations.
Warranties & Indemnification: Service quality warranties and indemnity clauses.
Insurance Requirements: Professional liability or other coverages.
Termination & Transition: Notice periods, wind-down services, and return of property.
Governing Law & Dispute Resolution: Applicable law and forum.
Signatures: Authorized representatives of both parties.`;
            break;
        case 'SLA':
            specificRequirements = `14. Service Level Agreement (SLA) Structure :
Introduction: Describes the objective of the SLA, which is to define expectations for service quality and performance.
Parties Involved: The agreement is between ${firstParty} and ${counterparty}, with full legal details such as registration number and address.
Parties and Scope: Includes details of the parties involved and a description of the services covered (e.g., software uptime guarantees).
Service Level Commitments: Detailed metrics (e.g., 99.9% uptime) and response times for issues.
Penalties for Non-Compliance: Escalating penalties for service failures.
Termination Clause: Includes conditions for terminating the SLA, specifying Termination for Cause and Termination for Convenience with defined notice periods.
ESG Clause (Environmental, Social, and Governance): Include a commitment to ESG principles in the service delivery, including sustainability targets and governance practices.
Signatures: Space for both parties’ representatives to sign.`;
            break;
        case 'FSA':
            specificRequirements = `15. Framework Service Agreement (FSA) Structure :
Introduction: This section sets the foundation for an ongoing service delivery relationship between ${firstParty} and the client. The FSA outlines the terms under which services will be provided over a set period, often for multiple projects or service engagements. It defines the general terms, including the rights and obligations of both parties, and serves as the overarching agreement that governs individual Statements of Work (SOW) or project-specific agreements.
Parties and Objectives: This section provides details on both ${firstParty} and the client. It includes the official names and addresses of the contracting parties, as well as a description of the scope of services covered under the framework. This could include ongoing support, maintenance, consulting, or other service offerings relevant to the [industry] (e.g., software development, solar panel services). The objectives should clarify the mutual intentions of both parties in entering this agreement.
Key Terms and Conditions: This section covers payment terms, including billing schedules, milestones, and any performance-linked payment mechanisms.  It should also detail renewal processes, including conditions for the extension of the agreement, the automatic renewal clause (if applicable), and procedures for amending or terminating the agreement after the initial term.
 Any conditions for project-specific agreements under the framework should be outlined here. This includes how each individual project or service request will be structured, including Statement of Work (SOW) issuance, deliverables, and agreed timelines.
Risk Management and Liability: This section should specify limitations of liability for both parties, addressing the maximum amount either party can be held liable for in the event of a breach or other failure to perform under the agreement.  Include risk-sharing mechanisms, such as indemnification clauses, ensuring that each party is protected from legal or financial consequences resulting from third-party claims.  Additionally, outline any exclusions to liability, particularly where force majeure events may limit or suspend performance.
ESG Clause (Environmental, Social, and Governance): Include provisions for compliance with Environmental, Social, and Governance (ESG) standards. This is particularly important in industries such as solar panel services and software. The clause should address environmental responsibility, social impact (e.g., diversity and inclusion), and governance practices, committing both parties to uphold certain sustainability standards and compliance with local laws (e.g., environmental laws, fair labor practices).
Signatures: This section provides space for the authorized representatives of both parties to sign, confirming their acceptance and understanding of the terms and conditions set forth in the agreement. The signature lines should include the name, title, and date of the signatories, ensuring both parties are legally bound by the agreement.`;
            break;
        case 'Termination':
            specificRequirements = `16. Termination Document :
Introduction: This section formalizes the termination of an existing agreement. It outlines the reasons for termination, such as breach of contract, completion of project scope, or mutual consent. It provides a clear framework for how the relationship between the parties will end, ensuring that both parties are aware of their obligations and rights upon termination.
Parties and Background: Specifies the agreement being terminated and identifies the parties involved. This section should provide a concise background on the original agreement, including any relevant context, milestones, and activities covered under that agreement, before the termination is effected.
Termination Terms: This section outlines the timeline for termination, including any necessary notice periods, and sets the conditions for either party to terminate the agreement.  The section should detail any obligations post-termination, such as returning intellectual property, settling outstanding invoices, and complying with confidentiality clauses that extend beyond the agreement's end date.  Additionally, any penalties for early termination (if applicable) should be explained, particularly if termination occurs before the agreed term ends or if there is a breach of contract.
Signatures: This section provides space for authorized representatives of both parties to sign, confirming their agreement to the terms of termination. The signature lines should include the name, title, and date of the signatories, ensuring both parties are legally bound by the termination terms.`;
            break;
        default:
            specificRequirements = `General Document Structure:
  Introduction: Provide a brief overview of the document's purpose and the parties involved.
  Parties and Background: Identify the parties involved and provide relevant background information.
  Key Terms and Conditions: Outline the main terms and conditions applicable to the document.
  Obligations of the Parties: Detail the responsibilities and obligations of each party.
  Termination Clause: Specify the conditions under which the document can be terminated.
  Signatures: Include space for the authorized representatives of both parties to sign, confirming their agreement to the terms.`;
            break;
    }
    return specificRequirements;
}

function getPromptOutputFormat(firstParty, counterparty) {

    const outputFormat = `Output Format
The output must be formatted as valid HTML, using markup such as h1, h2, sections or bullet points. Each document must be formatted in Arial, font size 11 or 12, with clearly structured headings (e.g., level 1, level 2, etc.). Missing sections or necessary content must be added depending on the document type. Ensure that all sections are appropriately detailed, maintaining the same quality and level of completeness, in line with legal best practices. Sections must follow an order based on market usage, ensuring that the document adheres to the typical structure of its type (e.g., agreements, contracts, amendments) as commonly practiced in the industry. 

Example Output:

Master Service Agreement 

This Master Service Agreement ("Agreement") is made effective as of November 22, 2022 (the "Effective Date") by and between ${firstParty} ("Company"), and ${counterparty} ("Service Provider").


1. Services 
Service Provider agrees to provide the coffee distibution services (the "Services") as outlined in one or more statements of work ("SOW") to be separately executed by the parties.

2. Provision of Supplies
If selected, Service Provider will also be responsible for providing all necessary cleaning supplies and equipment. Costs for these supplies will be reimbursed by Company as per the agreement in each SOW.

...

Ensure the HTML formatting is clean and adheres to standard practices. If possible, ensure spacing between paragraphs, and add bold for headlines.`;

    return outputFormat;
}

/**
 * Takes a generated data object, creates a prompt, calls the AI, 
 * and saves the resulting file to the specified subfolder.
 * @param {Object} docData The object of data generated by a ...Row function.
 * @param {GoogleAppsScript.Drive.Folder} subfolder The Google Drive folder to save the new file in.
 */
function processAndCreateFile(docData, subfolder) {
    if (!docData) {
        Logger.log("processAndCreateFile skipped because docData was null.");
        return;
    }

    // Destructure all needed properties from the data object
    const {
        agreementType, language, firstParty, counterparty,
        contractNumber, // The MSA's own number
        parentMsaContractNumber // The number from the parent MSA for SOWs/COs
    } = docData;

    const role = 'This GPT is designated to generate realistic sample agreements for use during AI demonstrations. It is tailored to create agreements with specific legal language and conditions that can be analyzed to return structured information.';
    const prompt = createPrompt(docData);

    try {
        const responseFromOpenAI = PreSalesOpenAI.executePrompt4o(role, prompt);

        // --- REFINEMENT: Determine the filename number based on document type ---
        let numberForFileName;
        if (agreementType.includes("MSA")) {
            // If it's an MSA, use its own specific contract number.
            numberForFileName = contractNumber;
        } else {
            // For all other documents (SOWs, COs, NDAs), generate a new, unique number.
            numberForFileName = generateContractNumber();
        }

        // Pass the correctly determined number to the file creation function.
        const newFileId = createFileInDriveV3(responseFromOpenAI, agreementType, language, numberForFileName);

        const newFile = DriveApp.getFileById(newFileId);
        newFile.moveTo(subfolder);
        newFile.setDescription(`Template for ${firstParty} and ${counterparty}`);

    } catch (error) {
        Logger.log(`Failed to create document for ${agreementType} with ${counterparty}. Original Error: ${error.message}`);
        // This includes the original error details in the message shown to the user.
        throw new Error(`Failed for ${agreementType}. Details: ${error.message}`);
    }
}

/**
 * Extracts individual document types from a formatted string.
 * @param {string} docTypeString The string, e.g., "Standard (NDA, MSA)".
 * @return {Array<string>} An array of types, e.g., ["NDA", "MSA"].
 */
function extractDocTypes(docTypeString) {
    const individualDocTypes = [];
    if (!docTypeString) {
        return individualDocTypes;
    }

    const categoryRegex = /\(([^)]+)\)/g;
    let match;
    while ((match = categoryRegex.exec(docTypeString)) !== null) {
        const docsInCategory = match[1].split(",").map(type => type.trim());
        individualDocTypes.push(...docsInCategory);
    }
    return individualDocTypes;
}

function generateContractNumber() {
    const randomPart = Math.floor(Math.random() * 100000).toString().padStart(5, '0'); // Random 5-digit number
    const contractNumber = `CN-${randomPart}`;
    return contractNumber;
}

function createFileInDriveV3(html, agreementType, language, contractNumber) {
    const languageAbbreviations = {
        Spanish: "[ES]",
        French: "[FR]",
        German: "[DE]",
        "Portuguese (PT)": "[PT]",
        "Portuguese (BR)": "[BR]",
        Japanese: "[JA]",
    };

    const langPrefix = languageAbbreviations[language] || "";

    // Use the provided contract number, or generate a new one if it doesn't exist.
    const finalContractNumber = contractNumber || generateContractNumber();

    // Clean filename: avoid leading spaces
    const fileNameParts = [];
    if (langPrefix) fileNameParts.push(langPrefix);
    fileNameParts.push(agreementType, finalContractNumber);
    const fileName = fileNameParts.join(" - ").trim();

    // Sanitize HTML content
    html = sanitizeHtml(html);

    // Create blob and file
    const blob = Utilities.newBlob(html, MimeType.HTML, `${fileName}.html`);
    const file = DriveApp.createFile(blob);

    // Convert to Google Doc
    const docFile = Drive.Files.copy(
        {
            title: fileName,
            mimeType: MimeType.GOOGLE_DOCS,
        },
        file.getId()
    );

    file.setTrashed(true);

    Logger.log("Google Docs File ID: " + docFile.id);
    return docFile.id;
}

// Proper sanitization function
function sanitizeHtml(html) {
    // Strip out any leftover code block tags
    html = html.replace(/```html|```/g, "").trim();

    // Escape your special markers once, globally, instead of brute-forcing 100 loops
    html = html.replace(/<# </g, "&lt;# &lt;").replace(/> #>/g, "&gt; #&gt;");

    return html;
}

function sendSlackNotification(email, agreementType, language, fileUrl) {
    console.log("Email" + email);
    const languageAbbreviations = {
        Spanish: "[ES]",
        French: "[FR]",
        German: "[DE]",
        "Portuguese (PT)": "[PT]",
        "Portuguese (BR)": "[BR]",
        Japanese: "[JA]",
    };

    // Add language abbreviation if not English
    const langPrefix = languageAbbreviations[language] || ""; // Default to an empty string for English or unknown languages

    const properties = PropertiesService.getScriptProperties();
    const url = properties.getProperty('SLACK_WEBHOOK_URL');

    if (!url) {
        Logger.log("SLACK_WEBHOOK_URL not set in Script Properties.");
        return;
    }
    const formData = {
        "submitter": email,
        "link": langPrefix + " Sample " + agreementType + " " + fileUrl
    };
    const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(formData)
    };

    try {
        UrlFetchApp.fetch(url, options);
    } catch (error) {
        Logger.log("Error sending Slack notification: " + error.message);
    }
}
