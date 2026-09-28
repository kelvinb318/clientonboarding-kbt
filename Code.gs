/************************************************************
 * CLIENT ONBOARDING PORTAL
 * GOOGLE APPS SCRIPT BACKEND
 *
 * Frontend:
 *   Netlify index.html
 *
 * Backend:
 *   Google Apps Script
 *
 * Services:
 *   Google Sheets
 *   Google Drive
 *   Google Docs
 *   MailApp
 *   OpenAI API
 ************************************************************/


/************************************************************
 * CONFIGURATION
 ************************************************************/

const SPREADSHEET_ID =
  "11ZM1dKxBoDXsRPT5bqdsGN4kTOOhZ024PJnsyqz9I8Q";

const SHEET_NAME =
  "Form Responses 1";


/*
 * OpenAI
 */
const OPENAI_API_URL =
  "https://api.openai.com/v1/chat/completions";

const OPENAI_MODEL =
  "gpt-4o-mini";


/*
 * Google Drive
 *
 * If this is left blank, the script will find or create
 * a folder named "Client Onboarding".
 *
 * If you already know the folder ID, place it here.
 */
const CLIENT_PARENT_FOLDER_ID = "";


/*
 * Default Google Drive folder name.
 */
const CLIENT_PARENT_FOLDER_NAME =
  "Client Onboarding";


/************************************************************
 * GET
 *
 * Used to test whether the deployed Web App is running.
 ************************************************************/

function doGet() {

  return jsonResponse({

    success: true,

    message:
      "Client Onboarding API is running."

  });

}


/************************************************************
 * POST
 *
 * Receives onboarding information from Netlify.
 ************************************************************/

function doPost(e) {

  try {

    /******************************************************
     * Validate POST request
     ******************************************************/

    if (
      !e ||
      !e.postData ||
      !e.postData.contents
    ) {

      return jsonResponse({

        success: false,

        message:
          "No POST data was received."

      });

    }


    /******************************************************
     * Parse JSON
     ******************************************************/

    var data;

    try {

      data =
        JSON.parse(
          e.postData.contents
        );

    } catch (parseError) {

      return jsonResponse({

        success: false,

        message:
          "The submitted data was not valid JSON."

      });

    }


    /******************************************************
     * Read submitted values
     ******************************************************/

    var firstName =
      cleanValue(data.firstName);

    var lastName =
      cleanValue(data.lastName);

    var email =
      cleanValue(data.email);

    var companyName =
      cleanValue(data.companyName);

    var projectOverview =
      cleanValue(data.projectOverview);

    var learningStyle =
      cleanValue(data.learningStyle);


    /******************************************************
     * Validate required fields
     ******************************************************/

    if (
      !firstName ||
      !lastName ||
      !email ||
      !companyName ||
      !projectOverview ||
      !learningStyle
    ) {

      return jsonResponse({

        success: false,

        message:
          "All required fields must be completed."

      });

    }


    /******************************************************
     * Validate email
     ******************************************************/

    if (!isValidEmail(email)) {

      return jsonResponse({

        success: false,

        message:
          "Please provide a valid email address."

      });

    }


    /******************************************************
     * Validate learning style
     ******************************************************/

    var validLearningStyles = [

      "Visual",

      "Auditory",

      "Reading/Writing",

      "Kinesthetic"

    ];


    if (
      validLearningStyles.indexOf(
        learningStyle
      ) === -1
    ) {

      return jsonResponse({

        success: false,

        message:
          "Invalid learning style selected."

      });

    }


    /******************************************************
     * Open Google Sheet
     ******************************************************/

    var spreadsheet =
      SpreadsheetApp.openById(
        SPREADSHEET_ID
      );


    var sheet =
      spreadsheet.getSheetByName(
        SHEET_NAME
      );


    if (!sheet) {

      throw new Error(
        'The sheet "' +
        SHEET_NAME +
        '" could not be found.'
      );

    }


    /******************************************************
     * AI CLASSIFICATION
     ******************************************************/

    var aiResult =
      classifyLearningStyle(

        learningStyle,

        projectOverview

      );


    /******************************************************
     * SAVE TO GOOGLE SHEET
     *
     * A Timestamp
     * B First Name
     * C Last Name
     * D Email
     * E Company Name
     * F Project Overview
     * G Learning Style
     * H AI Classification
     * I AI Reasoning
     * J Status
     * K Notes
     ******************************************************/

    sheet.appendRow([

      new Date(),

      firstName,

      lastName,

      email,

      companyName,

      projectOverview,

      learningStyle,

      aiResult.classification,

      aiResult.reasoning,

      "New",

      ""

    ]);


    /******************************************************
     * CREATE GOOGLE DRIVE FILE
     ******************************************************/

    var clientFile =
      createClientFile({

        firstName:
          firstName,

        lastName:
          lastName,

        email:
          email,

        companyName:
          companyName,

        projectOverview:
          projectOverview,

        learningStyle:
          learningStyle,

        aiClassification:
          aiResult.classification,

        aiReasoning:
          aiResult.reasoning

      });


    /******************************************************
     * SEND WELCOME EMAIL
     ******************************************************/

    var emailResult =
      sendWelcomeEmail({

        firstName:
          firstName,

        lastName:
          lastName,

        email:
          email,

        companyName:
          companyName,

        fileUrl:
          clientFile.fileUrl

      });


    /******************************************************
     * SUCCESS RESPONSE
     ******************************************************/

    return jsonResponse({

      success: true,

      message:
        "Client onboarding completed successfully.",

      clientName:
        firstName +
        " " +
        lastName,

      clientFileCreated:
        clientFile.success,

      clientFileUrl:
        clientFile.fileUrl,

      clientFolderUrl:
        clientFile.folderUrl,

      welcomeEmailSent:
        emailResult.success,

      welcomeEmailRecipient:
        email,

      fileError:
        clientFile.error || "",

      emailError:
        emailResult.error || ""

    });


  } catch (error) {

    console.error(
      "doPost ERROR:",
      error
    );


    return jsonResponse({

      success: false,

      message:
        error.message ||
        "An unexpected server error occurred."

    });

  }

}


/************************************************************
 * CREATE GOOGLE DRIVE FILE
 *
 * Creates ONE standalone Google Doc per new user.
 *
 * Structure:
 *
 * Client Onboarding/
 *
 *    John Smith - Onboarding Profile - timestamp
 *
 *    Jane Doe - Onboarding Profile - timestamp
 *
 *    Robert Jones - Onboarding Profile - timestamp
 *
 * There are NO individual user folders.
 ************************************************************/

function createClientFile(client) {

  try {

    /******************************************************
     * Find or create Client Onboarding folder
     ******************************************************/

    var parentFolder;


    if (CLIENT_PARENT_FOLDER_ID) {

      parentFolder =
        DriveApp.getFolderById(
          CLIENT_PARENT_FOLDER_ID
        );

    } else {

      var folders =
        DriveApp.getFoldersByName(
          CLIENT_PARENT_FOLDER_NAME
        );


      if (folders.hasNext()) {

        parentFolder =
          folders.next();

      } else {

        parentFolder =
          DriveApp.createFolder(
            CLIENT_PARENT_FOLDER_NAME
          );

      }

    }


    /******************************************************
     * Create unique file name
     ******************************************************/

    var fullName =
      client.firstName +
      " " +
      client.lastName;


    var timestamp =
      Utilities.formatDate(

        new Date(),

        Session.getScriptTimeZone(),

        "yyyy-MM-dd HH-mm-ss"

      );


    var documentName =
      fullName +
      " - Onboarding Profile - " +
      timestamp;


    /******************************************************
     * Create Google Doc
     ******************************************************/

    var document =
      DocumentApp.create(
        documentName
      );


    /******************************************************
     * Move document into Client Onboarding folder
     ******************************************************/

    var documentFile =
      DriveApp.getFileById(
        document.getId()
      );


    documentFile.moveTo(
      parentFolder
    );


    /******************************************************
     * Build document
     ******************************************************/

    var body =
      document.getBody();


    body.clear();


    /******************************************************
     * TITLE
     ******************************************************/

    body.appendParagraph(
      "CLIENT ONBOARDING PROFILE"
    )
    .setHeading(
      DocumentApp.ParagraphHeading.TITLE
    );


    body.appendParagraph(
      "Generated by Client Onboarding Portal"
    );


    body.appendParagraph(
      ""
    );


    /******************************************************
     * CLIENT INFORMATION
     ******************************************************/

    body.appendParagraph(
      "CLIENT INFORMATION"
    )
    .setHeading(
      DocumentApp.ParagraphHeading.HEADING1
    );


    body.appendParagraph(
      "First Name: " +
      client.firstName
    );


    body.appendParagraph(
      "Last Name: " +
      client.lastName
    );


    body.appendParagraph(
      "Email: " +
      client.email
    );


    body.appendParagraph(
      "Company: " +
      client.companyName
    );


    body.appendParagraph(
      ""
    );


    /******************************************************
     * PROJECT INFORMATION
     ******************************************************/

    body.appendParagraph(
      "PROJECT OVERVIEW"
    )
    .setHeading(
      DocumentApp.ParagraphHeading.HEADING1
    );


    body.appendParagraph(
      client.projectOverview
    );


    body.appendParagraph(
      ""
    );


    /******************************************************
     * LEARNING PREFERENCES
     ******************************************************/

    body.appendParagraph(
      "LEARNING PREFERENCES"
    )
    .setHeading(
      DocumentApp.ParagraphHeading.HEADING1
    );


    body.appendParagraph(
      "Selected Learning Style: " +
      client.learningStyle
    );


    body.appendParagraph(
      "AI Classification: " +
      client.aiClassification
    );


    body.appendParagraph(
      ""
    );


    /******************************************************
     * AI REASONING
     ******************************************************/

    body.appendParagraph(
      "AI REASONING"
    )
    .setHeading(
      DocumentApp.ParagraphHeading.HEADING1
    );


    body.appendParagraph(
      client.aiReasoning
    );


    body.appendParagraph(
      ""
    );


    /******************************************************
     * STATUS
     ******************************************************/

    body.appendParagraph(
      "ONBOARDING STATUS"
    )
    .setHeading(
      DocumentApp.ParagraphHeading.HEADING1
    );


    body.appendParagraph(
      "Status: New"
    );


    body.appendParagraph(
      "Created: " +
      new Date().toString()
    );


    /******************************************************
     * Save document
     ******************************************************/

    document.saveAndClose();


    /******************************************************
     * Return file information
     ******************************************************/

    return {

      success: true,

      fileId:
        document.getId(),

      fileUrl:
        document.getUrl(),

      folderUrl:
        parentFolder.getUrl(),

      error:
        ""

    };


  } catch (error) {

    console.error(
      "createClientFile ERROR:",
      error
    );


    return {

      success: false,

      fileId:
        "",

      fileUrl:
        "",

      folderUrl:
        "",

      error:
        error.message ||
        "Unable to create Google Drive file."

    };

  }

}


/************************************************************
 * SEND WELCOME EMAIL
 ************************************************************/

function sendWelcomeEmail(client) {

  try {

    var subject =
      "Welcome to the Client Onboarding Portal";


    /******************************************************
     * Plain text email
     ******************************************************/

    var plainTextBody =

      "Hello " +
      client.firstName +
      ",\n\n" +

      "Welcome to the Client Onboarding Portal.\n\n" +

      "Your onboarding information has been " +
      "successfully received.\n\n" +

      "Company: " +
      client.companyName +
      "\n\n";


    if (client.fileUrl) {

      plainTextBody +=

        "Your onboarding profile has been created " +
        "and is available here:\n\n" +

        client.fileUrl +
        "\n\n";

    }


    plainTextBody +=

      "Thank you,\n" +
      "Client Onboarding Team";


    /******************************************************
     * HTML email
     ******************************************************/

    var htmlBody =

      "<div style=\"" +

      "font-family:Arial,Helvetica,sans-serif;" +

      "line-height:1.6;" +

      "color:#333333;" +

      "max-width:650px;" +

      "margin:auto;" +

      "\">" +

      "<h2 style=\"color:#2563eb;\">" +

      "Welcome, " +

      escapeHtml(
        client.firstName
      ) +

      "!" +

      "</h2>" +

      "<p>" +

      "Welcome to the Client Onboarding Portal." +

      "</p>" +

      "<p>" +

      "Your onboarding information has been " +
      "successfully received." +

      "</p>" +

      "<p>" +

      "<strong>Company:</strong> " +

      escapeHtml(
        client.companyName
      ) +

      "</p>";


    if (client.fileUrl) {

      htmlBody +=

        "<p>" +

        "Your onboarding profile has been " +
        "created and is available here:" +

        "</p>" +

        "<p>" +

        "<a href=\"" +

        client.fileUrl +

        "\" target=\"_blank\" " +

        "style=\"" +

        "display:inline-block;" +

        "padding:12px 18px;" +

        "background:#2563eb;" +

        "color:#ffffff;" +

        "text-decoration:none;" +

        "border-radius:6px;" +

        "font-weight:bold;" +

        "\">" +

        "View Your Onboarding Profile" +

        "</a>" +

        "</p>";

    }


    htmlBody +=

      "<p>" +

      "Thank you,<br>" +

      "Client Onboarding Team" +

      "</p>" +

      "</div>";


    /******************************************************
     * Send email
     ******************************************************/

    MailApp.sendEmail({

      to:
        client.email,

      subject:
        subject,

      body:
        plainTextBody,

      htmlBody:
        htmlBody

    });


    return {

      success:
        true,

      error:
        ""

    };


  } catch (error) {

    console.error(
      "sendWelcomeEmail ERROR:",
      error
    );


    return {

      success:
        false,

      error:
        error.message ||
        "Unable to send welcome email."

    };

  }

}


/************************************************************
 * OPENAI CLASSIFICATION
 ************************************************************/

function classifyLearningStyle(
  learningStyle,
  projectOverview
) {

  try {

    /******************************************************
     * Get API key from Script Properties
     ******************************************************/

    var apiKey =
      PropertiesService
        .getScriptProperties()
        .getProperty(
          "OPENAI_API_KEY"
        );


    if (!apiKey) {

      throw new Error(
        "OPENAI_API_KEY is not configured in Apps Script Script Properties."
      );

    }


    /******************************************************
     * System prompt
     ******************************************************/

    var systemPrompt =

      "You are an onboarding assistant. " +

      "Analyze the client's selected learning style " +

      "and project overview. " +

      "Classify the client into exactly one of " +

      "these four categories: " +

      "Visual, Auditory, Reading/Writing, or Kinesthetic. " +

      "Return ONLY valid JSON in this exact structure: " +

      "{ " +

      "\"classification\": \"Visual\", " +

      "\"reasoning\": \"Brief explanation\" " +

      "}";


    /******************************************************
     * User prompt
     ******************************************************/

    var userPrompt =

      "Selected learning style: " +

      learningStyle +

      "\n\n" +

      "Project overview: " +

      projectOverview;


    /******************************************************
     * OpenAI request
     ******************************************************/

    var payload = {

      model:
        OPENAI_MODEL,

      messages: [

        {

          role:
            "system",

          content:
            systemPrompt

        },

        {

          role:
            "user",

          content:
            userPrompt

        }

      ],

      temperature:
        0.2

    };


    var options = {

      method:
        "post",

      contentType:
        "application/json",

      headers: {

        Authorization:
          "Bearer " +
          apiKey

      },

      payload:
        JSON.stringify(
          payload
        ),

      muteHttpExceptions:
        true

    };


    var response =
      UrlFetchApp.fetch(

        OPENAI_API_URL,

        options

      );


    var responseCode =
      response.getResponseCode();


    var responseText =
      response.getContentText();


    /******************************************************
     * Check OpenAI response
     ******************************************************/

    if (
      responseCode < 200 ||
      responseCode >= 300
    ) {

      throw new Error(

        "OpenAI API error (" +

        responseCode +

        "): " +

        responseText

      );

    }


    /******************************************************
     * Parse OpenAI response
     ******************************************************/

    var openAIResponse =
      JSON.parse(
        responseText
      );


    if (
      !openAIResponse.choices ||
      !openAIResponse.choices.length
    ) {

      throw new Error(
        "OpenAI returned no classification."
      );

    }


    var content =
      openAIResponse
        .choices[0]
        .message
        .content
        .trim();


    /******************************************************
     * Remove Markdown code fences
     ******************************************************/

    content =
      content
        .replace(
          /^```json\s*/i,
          ""
        )
        .replace(
          /^```\s*/i,
          ""
        )
        .replace(
          /\s*```$/i,
          ""
        )
        .trim();


    /******************************************************
     * Parse AI JSON
     ******************************************************/

    var result =
      JSON.parse(
        content
      );


    /******************************************************
     * Validate classification
     ******************************************************/

    var validClassifications = [

      "Visual",

      "Auditory",

      "Reading/Writing",

      "Kinesthetic"

    ];


    if (
      validClassifications.indexOf(
        result.classification
      ) === -1
    ) {

      throw new Error(
        "OpenAI returned an invalid learning-style classification."
      );

    }


    return {

      classification:
        result.classification,

      reasoning:
        result.reasoning || ""

    };


  } catch (error) {

    console.error(
      "classifyLearningStyle ERROR:",
      error
    );


    /*
     * Keep onboarding operational if OpenAI
     * is temporarily unavailable.
     */

    return {

      classification:
        learningStyle,

      reasoning:
        "The submitted learning style was retained. " +
        "AI classification was unavailable: " +
        error.message

    };

  }

}


/************************************************************
 * EMAIL VALIDATION
 ************************************************************/

function isValidEmail(email) {

  var emailPattern =
    /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  return emailPattern.test(
    email
  );

}


/************************************************************
 * CLEAN VALUES
 ************************************************************/

function cleanValue(value) {

  if (
    value === null ||
    value === undefined
  ) {

    return "";

  }


  return String(
    value
  ).trim();

}


/************************************************************
 * HTML ESCAPING
 ************************************************************/

function escapeHtml(value) {

  if (
    value === null ||
    value === undefined
  ) {

    return "";

  }


  return String(value)

    .replace(
      /&/g,
      "&amp;"
    )

    .replace(
      /</g,
      "&lt;"
    )

    .replace(
      />/g,
      "&gt;"
    )

    .replace(
      /"/g,
      "&quot;"
    )

    .replace(
      /'/g,
      "&#039;"
    );

}


/************************************************************
 * JSON RESPONSE
 *
 * IMPORTANT:
 * Returns JSON as TEXT so the Netlify frontend can
 * explicitly read and parse the response.
 ************************************************************/

function jsonResponse(data) {

  return ContentService

    .createTextOutput(
      JSON.stringify(data)
    )

    .setMimeType(
      ContentService.MimeType.TEXT
    );

}
