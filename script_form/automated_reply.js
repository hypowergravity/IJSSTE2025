function sendAutomatedReply(e) {
  try {
    // Get all the form responses
    var itemResponses = e.response.getItemResponses();
    
    // Extract information from the form
    var email = "";
    var firstName = "";
    var lastName = "";
    var institution = "";
    var category = "";
    
    // Loop through all responses to find each answer
    for (var i = 0; i < itemResponses.length; i++) {
      var itemResponse = itemResponses[i];
      var question = itemResponse.getItem().getTitle();
      var answer = itemResponse.getResponse();
      
      // Use flexible matching to handle your form's question structure
      // This handles "Personal Information\nFirst Name" or just "First Name"
      if (question.match(/First Name/i)) {
        firstName = answer.trim();
      } else if (question.match(/Last Name/i)) {
        lastName = answer.trim();
      } else if (question.match(/Email Address/i)) {
        email = answer.trim();
      } else if (question.match(/Institution|Organization/i)) {
        institution = answer.trim();
      } else if (question.match(/Participant Category/i)) {
        category = answer;
      }
    }
    
    // Check if we got an email address
    if (!email || email === "") {
      Logger.log("No email address found in form submission");
      return;
    }
    
    // Validate that the email looks correct
    if (!email.includes("@") || !email.includes(".")) {
      Logger.log("Invalid email address: " + email);
      return;
    }
    
    // Check if we got a first name, use a generic greeting if not
    var greeting = firstName ? firstName : "Participant";
    
    // Get the submission timestamp
    var timestamp = e.response.getTimestamp();
    var formattedTime = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "MMMM dd, yyyy 'at' hh:mm a");
    
    // Email subject - personalized with first name
    var subject = firstName 
      ? "Welcome " + firstName + " - IJSSTE Hokkaido Symposium 2025 Registration Confirmed"
      : "Registration Confirmation - IJSSTE Hokkaido Symposium 2025";
    
    // Build the HTML email with personal touches throughout
    var htmlMessage = "<html><body style='font-family: Arial, sans-serif; color: #333; line-height: 1.6;'>";
    htmlMessage += "<div style='max-width: 650px; margin: 0 auto; padding: 20px;'>";
    
    // Header
    htmlMessage += "<h2 style='color: #2c5aa0; border-bottom: 3px solid #2c5aa0; padding-bottom: 10px;'>IJSSTE Hokkaido Symposium 2025</h2>";
    htmlMessage += "<h3 style='color: #555; font-style: italic; margin-top: 5px;'>AI and Experiment: Shaping the Next Generation of Molecular and Materials Design</h3>";
    
    // Personal greeting
    htmlMessage += "<p style='margin-top: 30px; font-size: 16px;'>Dear " + greeting + ",</p>";
    
    // Personal confirmation message
    htmlMessage += "<p>Thank you for registering for the <strong>IJSSTE Hokkaido Symposium 2025</strong>! ";
    htmlMessage += "We're excited to have you join us for this exciting symposium on molecular and materials design.</p>";
    
    htmlMessage += "<p>Your registration was successfully submitted on <strong>" + formattedTime + "</strong>.</p>";
    
    // Registration details box
    htmlMessage += "<div style='background-color: #f8f9fa; padding: 20px; border-left: 4px solid #2c5aa0; margin: 25px 0;'>";
    htmlMessage += "<h4 style='margin-top: 0; color: #2c5aa0;'>Your Registration Details</h4>";
    
    htmlMessage += "<table style='width: 100%; border-collapse: collapse;'>";
    htmlMessage += "<tr><td style='padding: 8px 0; width: 40%;'><strong>Name:</strong></td><td style='padding: 8px 0;'>" + firstName + " " + lastName + "</td></tr>";
    htmlMessage += "<tr><td style='padding: 8px 0;'><strong>Email:</strong></td><td style='padding: 8px 0;'>" + email + "</td></tr>";
    htmlMessage += "<tr><td style='padding: 8px 0;'><strong>Institution:</strong></td><td style='padding: 8px 0;'>" + institution + "</td></tr>";
    htmlMessage += "<tr><td style='padding: 8px 0;'><strong>Category:</strong></td><td style='padding: 8px 0;'>" + category + "</td></tr>";
    htmlMessage += "<tr><td style='padding: 8px 0;'><strong>Registration Date:</strong></td><td style='padding: 8px 0;'>" + formattedTime + "</td></tr>";
    htmlMessage += "</table>";
    
    htmlMessage += "</div>";
    
    // Important notes for oral/poster presenters - personalized with slot information
    htmlMessage += "<div style='background-color: #fff3cd; padding: 15px; border-radius: 5px; margin: 20px 0;'>";
    htmlMessage += "<p style='margin: 0;'><strong>📌 Important Information for Presenters:</strong></p>";
    htmlMessage += "<p style='margin: 10px 0 0 0;'>";
    if (firstName) {
      htmlMessage += greeting + ", if you uploaded the abstract for poster or oral presentation, we will get back to you ";
    } 
    htmlMessage += "For further communication please contact <a href='mailto:ijsste.hokkaido.symposium.2025@gmail.com' style='color: #2c5aa0;'>ijsste.hokkaido.symposium.2025@gmail.com</a></p>";
    
    // Add the slot limitation information
    htmlMessage += "<p style='margin: 15px 0 0 0; padding: 10px; background-color: #fff; border-left: 3px solid #ff9800;'>";
    htmlMessage += "<strong>Please Note:</strong> There are limited slots available for oral presentations. If slots are unavailable, we offer our sincere apologies, and your submission will be considered for poster presentation instead.";
    htmlMessage += "</p>";

    htmlMessage += "<p style='margin: 15px 0 0 0; padding: 10px; background-color: #e8f4f8; border-left: 3px solid #2c5aa0;'>";
    htmlMessage += "<strong>Please Note:</strong> The registration fees will be collected on the event day. Registration includes Lunch box and Tea/Coffee.";
    htmlMessage += "</p>";
    
    htmlMessage += "</div>";
    
    // Next steps - personalized
    htmlMessage += "<h4 style='color: #2c5aa0; margin-top: 30px;'>What's Next, " + greeting + "?</h4>";
    htmlMessage += "<ul style='line-height: 1.8;'>";
    htmlMessage += "<li>We'll send you detailed information about the symposium schedule as we get closer to the event date.</li>";
    htmlMessage += "<li>If you submitted an abstract for oral presentation, we will notify you about the presentation format (oral or poster) after reviewing all submissions.</li>";
    htmlMessage += "<li>Please keep this email for your records - you may need it for reference.</li>";
    htmlMessage += "<li>If you have any questions or need to update your registration details, please don't hesitate to contact us at <a href='mailto:ijsste.hokkaido.symposium.2025@gmail.com' style='color: #2c5aa0;'>ijsste.hokkaido.symposium.2025@gmail.com</a></li>";
    htmlMessage += "</ul>";
    
    // Personal closing
    htmlMessage += "<p style='margin-top: 30px;'>";
    if (firstName) {
      htmlMessage += greeting + ", we're looking forward to welcoming you to Hokkaido and to what promises to be an outstanding symposium!";
    } else {
      htmlMessage += "We're looking forward to welcoming you to Hokkaido and to what promises to be an outstanding symposium!";
    }
    htmlMessage += "</p>";
    
    htmlMessage += "<hr style='border: none; border-top: 1px solid #ddd; margin: 30px 0;'>";
    htmlMessage += "<p style='color: #666; font-size: 14px; margin-bottom: 5px;'>Warm regards,</p>";
    htmlMessage += "<p style='color: #666; font-size: 14px; margin-top: 5px;'><strong>IJSSTE Hokkaido Symposium 2025 Organizing Committee</strong></p>";
    
    htmlMessage += "</div>";
    htmlMessage += "</body></html>";
    
    // Send the HTML formatted email
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlMessage,
      name: "IJSSTE Hokkaido Symposium 2025"
    });
    
    Logger.log("Registration confirmation email sent successfully to: " + email + " (Name: " + firstName + " " + lastName + ")");
    
  } catch (error) {
    Logger.log("Error sending confirmation email: " + error.toString());
  }
}

function setupTrigger() {
  // Remove any existing triggers for this function to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendAutomatedReply') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create the new trigger for form submissions
  var form = FormApp.getActiveForm();
  ScriptApp.newTrigger('sendAutomatedReply')
    .forForm(form)
    .onFormSubmit()
    .create();
  
  Logger.log("Trigger created successfully! Registration confirmation emails are now active.");
}

function onOpen() {
  // Empty function to prevent onOpen errors
}