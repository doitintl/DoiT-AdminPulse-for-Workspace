function contactPartner() {
  try {
    // Compose the email
    var subject = "Assistance Request";
    var body = "Hi DoiT Team,\n\nSomeone has requested assistance with the Security Checklist for Workspace Admins. Please do your best to help them.";

    var recipient = "workspace-security@doit.com";

    // Send the email using MailApp service
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: body,
    });

    // Log confirmation
    Logger.log("Email Sent successfully.");
  } catch (error) {
    // Log any errors
    Logger.log("Error: " + error.toString());
  }
}
