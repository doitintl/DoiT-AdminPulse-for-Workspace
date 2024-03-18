function contactPartner() {
  try {
    // Get the user's email address
    var userEmail = Session.getActiveUser().getEmail();

    // Get the current document ID
    var docId = SpreadsheetApp.getActiveSpreadsheet().getId();

    // Generate the Google Sheet link
    var sheetLink = `https://docs.google.com/spreadsheets/d/${docId}/edit`;

    // Compose the email
    var subject = `Assistance Request - ${userEmail}`;
    var body = `Hi DoiT Team,\n\n${userEmail} has requested assistance with the Security Checklist for Workspace Admins. Please do your best to help them. \n\nA link to the customer's Security Checklist is included below: \n${sheetLink}`;

    var recipient = "workspace-security@doit.com";

    // Send the email using MailApp service
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: body,
    });

    // Display a message to the user
    var ui = SpreadsheetApp.getUi();
    ui.alert(`Email Sent! \n\nDoiT may request viewer access to assist you.`);
    
    // Log confirmation      
    console.log("Email Sent successfully.");
  } catch (error) {
    // Log any errors
    console.error("Error: " + error.toString());
  }
}
