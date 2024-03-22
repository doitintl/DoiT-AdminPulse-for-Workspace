function contactPartner() {
  try {
    // Constants
    const userEmail = Session.getUser().getEmail();
    const docId = SpreadsheetApp.getActiveSpreadsheet().getId();
    const sheetLink = `https://docs.google.com/spreadsheets/d/${docId}/edit`;
    const subject = `Assistance Request - ${userEmail}`;
    const body = `Hi DoiT Team,\n\n${userEmail} has requested assistance with the Security Checklist for Workspace Admins. Please do your best to help them. \n\nA link to the customer's Security Checklist is included below: \n${sheetLink}`;
    const recipient = "workspace-security@doit.com";

    // Send the email using MailApp service
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: body,
    });

    // Display a message to the user
    const ui = SpreadsheetApp.getUi();
    ui.alert(`Email Sent! \n\nDoiT may request viewer access to assist you.`);

    // Log confirmation
    console.log("Email Sent successfully.");
  } catch (error) {
    // Log any errors
    console.error("Error: " + error.toString());

    // Display troubleshooting suggestions
    const ui = SpreadsheetApp.getUi();
    if (
      error.message.includes(
        "Required permissions: https://www.googleapis.com/auth/userinfo.email",
      )
    ) {
      ui.alert(
        "Error: You do not have permission to access user information. Please enable the required permission and try again.",
      );
    } else {
      ui.alert(
        "Error sending email. Please check your internet connection and try again. If the problem persists, contact support at workspace-security@doit.com.",
      );
    }
  }
}
