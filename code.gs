// ======================================
// CODE.GS - PRODUCTION FIXED VERSION
// ======================================


// ======================================
// LOAD WEB APP
// ======================================
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("Medical Assistant Skills Checklist");
}


// ======================================
// INCLUDE FILES
// ======================================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ======================================
// GET BASE64 LOGO
// ======================================
function getBase64Image() {
  try {
    const imageUrl = "https://unizonhealth.com/data/nnlogo.png";
    const response = UrlFetchApp.fetch(imageUrl);
    const blob = response.getBlob();

    return "data:" +
      blob.getContentType() +
      ";base64," +
      Utilities.base64Encode(blob.getBytes());

  } catch (error) {
    Logger.log("Logo error: " + error);
    return "";
  }
}


// ======================================
// MAIN SUBMIT FUNCTION
// ======================================
function submitForm(data) {

  const lock = LockService.getScriptLock();

  try {

    Logger.log("Incoming Data: " + JSON.stringify(data));

    if (!data) {
      return {
        status: "error",
        message: "No form data received."
      };
    }

    lock.waitLock(30000);

    // ======================================
    // SHEET SAVE
    // ======================================
    const sheet = SpreadsheetApp
      .openById("1hYEZvHxrEAHARsEbFORZC2qge2Tx8iB44dHsJoTAatY")
      .getSheetByName("Sheet1");

    sheet.appendRow([
      new Date(),
      data.fullName || "",
      data.email || "",
      data.ssn || "",
      data.date || ""
    ]);

    // ======================================
    // TEMPLATE
    // ======================================
    const template = HtmlService.createTemplateFromFile("pdf");

    template.fullName = data.fullName || "";
    template.email = data.email || "";
    template.ssn = data.ssn || "";
    template.date = data.date || "";
    template.additionalSkills = data.additionalSkills || "";
    template.additionalTraining = data.additionalTraining || "";
    template.signature = data.signature || "";
    template.submittedOn = new Date().toLocaleString();
    template.logoBase64 = getBase64Image();

    // ======================================
    // SKILLS MAPPING
    // ======================================
    function mapSkills(names, prefix) {
      return names.map((name, i) => ({
        name: name,
        value: data[prefix + i] || ""
      }));
    }

    const generalNames = [
      "Awareness of HCAHPS",
      "Administrative procedures",
      "MDS Coordinator",
      "Admit/discharge patients",
      "Blood glucose monitoring",
      "Coordinate scheduling",
      "Documentation",
      "Dressing changes",
      "Familiarity with advanced directives",
      "HIPAA regulations",
      "Isolation techniques",
      "Observe for adverse medication reaction",
      "Alert licensed staff of medication reaction",
      "Patient education",
      "Position/transfer patients",
      "Prepare reports",
      "Pulse oximetry",
      "Lab draw",
      "Screen/direct provider calls",
      "Screen/direct patient calls",
      "Urine dipstick",
      "Universal precautions",
      "Vital signs",
      "Wound care"
    ];

    template.generalSkills = mapSkills(generalNames, "general");

    template.cardiacSkills = mapSkills([
      "Assist with emergency",
      "Perform 12 lead EKG",
      "Use of cardiac monitor"
    ], "cardiac");

    template.vascularSkills = mapSkills([
      "Apply/monitor noninvasive BP monitor",
      "Intake and output",
      "Discontinue peripheral IV",
      "Manual BP",
      "Phlebotomy draws"
    ], "vascular");

    template.pulmonarySkills = mapSkills([
      "Apply nasal cannula/face mask",
      "Incentive spirometry",
      "O2 saturation monitor"
    ], "pulmonary");

    template.neuroSkills = mapSkills([
      "Assist with lumbar puncture",
      "Neurological evaluation",
      "Seizure precautions"
    ], "neuro");

    template.orthoSkills = mapSkills([
      "Cast care",
      "Crutch walking",
      "Traction"
    ], "ortho");

    template.giSkills = mapSkills([
      "Assist with feeding",
      "Nutritional evaluation",
      "Instruct/obtain clean catch urine",
      "Straight/Foley catheter female",
      "Straight/Foley catheter male"
    ], "gi");

    template.medicationSkills = mapSkills([
      "Vitamins, minerals, herbs",
      "Antibiotics",
      "Antifungal",
      "Antiviral",
      "Psychotropic",
      "Ophthalmic medications",
      "Aural medications",
      "Respiratory system medications",
      "Cardiovascular system medications",
      "Gastrointestinal system medications",
      "Urinary system medications",
      "Reproductive system medications",
      "Endocrine system medications",
      "Musculoskeletal system medications",
      "Nervous system medications",
      "Immunizations",
      "Intramuscular (IM)",
      "Subcutaneous (SQ)",
      "Intradermal",
      "Z-track"
    ], "meds");

    template.ageSkills = mapSkills([
      "Newborn/neonate (birth-30 days)",
      "Infant (31 days-1 year)",
      "Toddler (ages 2-3 years)",
      "Preschool (ages 4-5 years)",
      "School age (ages 6-12 years)",
      "Adolescent (ages 13-21 years)",
      "Young adult (ages 22-39 years)",
      "Adult (ages 40-64 years)",
      "Older adult (ages 65-79 years)",
      "Elderly (ages 80+ years)"
    ], "age");

    // ======================================
    // CREATE HTML
    // ======================================
    const htmlOutput = template.evaluate();
    const html = htmlOutput.getContent();

    // ======================================
    // FIXED PDF GENERATION (STABLE)
    // ======================================
    const htmlBlob = Utilities.newBlob(html, MimeType.HTML, "file.html");

    const pdfBlob = htmlBlob.getAs(MimeType.PDF)
      .setName("Medical_Assistant_Checklist_" + (data.fullName || "User") + ".pdf");

    Logger.log("PDF generated successfully");

    // ======================================
    // EMAIL RECIPIENTS
    // ======================================
    const recipients = [
      "muneeb.unizon@gmail.com",
      "uddeshya.unizon@gmail.com",
      "haider.unizon@gmail.com",
      "lightson.unizon@gmail.com"
    ].join(",");

    Logger.log("Sending to internal: " + recipients);

    // ======================================
    // INTERNAL EMAIL
    // ======================================
    try {
      MailApp.sendEmail({
        to: recipients,
        subject: "New Medical Assistant Checklist Submission",
        body:
          "A new submission was received.\n\n" +
          "Name: " + (data.fullName || "") + "\n" +
          "Email: " + (data.email || "") + "\n" +
          "Date: " + (data.date || "") + "\n",
        attachments: [pdfBlob]
      });
      Logger.log("Internal email sent");
    } catch (e) {
      Logger.log("Internal email failed: " + e);
    }

    // ======================================
    // USER EMAIL
    // ======================================
    try {
      if (data.email && data.email.trim() !== "") {

        MailApp.sendEmail({
          to: data.email.trim(),
          subject: "Your Medical Assistant Skills Checklist",
          body:
            "Thank you for your submission.\n\nAttached is your PDF copy.",
          attachments: [pdfBlob]
        });

        Logger.log("User email sent");
      }
    } catch (e) {
      Logger.log("User email failed: " + e);
    }

    // ======================================
    // SUCCESS RESPONSE
    // ======================================
    return {
      status: "success",
      message: "Form submitted successfully."
    };

  } catch (error) {

    Logger.log("ERROR: " + error);

    return {
      status: "error",
      message: error.toString()
    };

  } finally {

    if (lock && lock.hasLock()) {
      lock.releaseLock();
    }
  }
}
// ======================================
// TEST SUBMIT FUNCTION (FOR DEBUGGING)
// ======================================
function testSubmitForm() {

  const sampleData = {
    fullName: "Test User",
    email: "testuser@example.com",
    ssn: "123-45-6789",
    date: new Date().toISOString().split("T")[0],

    additionalSkills: "Test skills",
    additionalTraining: "Test training",
    signature: "Test Signature",

    // sample checkbox values for skills (optional)
    general0: "Yes",
    general1: "No",
    cardiac0: "Yes",
    vascular0: "No",
    pulmonary0: "Yes"
  };

  Logger.log("Running testSubmitForm...");

  const result = submitForm(sampleData);

  Logger.log("Result: " + JSON.stringify(result));

  return result;
}