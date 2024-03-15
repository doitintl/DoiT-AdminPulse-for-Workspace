function contactPartner() {
  try {
    // Get the user's email address
    var userEmail = Session.getActiveUser().getEmail();

    // Compose the email
    var subject = `Assistance Request - ${userEmail}`;
    var body = `Hi DoiT Team,\n\n${userEmail} has requested assistance with the Security Checklist for Workspace Admins. Please do your best to help them.`;

    var recipient = "workspace-security@doit.com";

    // Send the email using MailApp service
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: body,
    });

    // Log confirmation
    console.log("Email Sent successfully.");
  } catch (error) {
    // Log any errors
    console.error("Error: " + error.toString());
  }
}
