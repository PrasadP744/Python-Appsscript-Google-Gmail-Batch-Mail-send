function sendCredentials() {
  try {
    // Use the correct Google Sheets ID
    var spreadsheetId = " Your spreadsheet ID";
    
    console.log("Opening spreadsheet with ID: " + spreadsheetId);
    
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheets()[0];
    
    console.log("Successfully opened sheet: " + sheet.getName());
    console.log("Sheet has " + sheet.getLastRow() + " rows and " + sheet.getLastColumn() + " columns");
    
    var data = sheet.getDataRange().getValues();
    
    console.log("Headers (Row 1):", data[0]);
    
    if (data.length < 2) {
      console.log("No data rows found (only headers)");
      return;
    }
    
    var successCount = 0;
    var errorCount = 0;
    
    for (var i = 1; i < data.length; i++) {
      // CORRECT column mapping based on your spreadsheet:
      // A=0: Sl.No, B=1: User Name, C=2: Login ID, D=3: Password, E=4: Email
      var slNo     = data[i][0];  // Column A - Serial Number
      var userName = data[i][1];  // Column B - User Name
      var loginId  = data[i][2];  // Column C - Login ID
      var password = data[i][3];  // Column D - Password
      var email    = data[i][4];  // Column E - Email

      console.log("Row " + (i+1) + " - Processing: " + userName);
      console.log("Email: " + email);

      // Skip empty rows or rows without email
      if (!email || email.toString().trim() === "") {
        console.log("Skipping row " + (i+1) + " - no email found");
        continue;
      }

      // Skip rows where email is obviously not an email (like password strings)
      var emailStr = email.toString().trim();
      var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      
      if (!emailRegex.test(emailStr)) {
        console.log("Skipping row " + (i+1) + " - invalid email format: " + emailStr);
        errorCount++;
        continue;
      }

      console.log("✅ Valid email found: " + emailStr);
      console.log("Preparing email for: " + userName + " (Login: " + loginId + ")");

      try {
        var subject = "🔐 Your VPN Access Credentials";
        var body = 
          "Dear " + userName + ",\n\n" +
          "Your VPN access credentials are ready:\n\n" +
          "🔑 Login ID: " + loginId + "\n" +
          "🔒 Password: " + password + "\n\n" +
          "⚠️ IMPORTANT SECURITY NOTES:\n" +
          "• Keep these credentials confidential\n" +
          "• Do not share with anyone\n" +
          "• Contact IT support if you have any issues\n\n" +
          "For VPN setup instructions, please refer to the IT documentation.\n\n" +
          "Best regards,\n" +
          "IT Security Team";

        // Send email
        MailApp.sendEmail({
          to: emailStr,
          subject: subject,
          body: body
        });
        
        console.log("✅ Email sent successfully to: " + emailStr);
        successCount++;
        
        // Add delay to avoid rate limiting
        Utilities.sleep(2000);
        
      } catch (emailError) {
        console.error("❌ Failed to send email to " + emailStr + ": " + emailError.toString());
        errorCount++;
      }
    }
    
    console.log("\n🏁 Process completed!");
    console.log("✅ Successfully sent: " + successCount + " emails");
    console.log("❌ Failed: " + errorCount + " emails");
    
  } catch (error) {
    console.error("💥 Error in sendCredentials function: " + error.toString());
  }
}

// Test function - run this first
function testSendToOneUser() {
  try {
    var spreadsheetId = "Your Spreadsheet ID";
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    
    console.log("Headers:", data[0]);
    
    if (data.length < 2) {
      console.log("No user data found");
      return;
    }
    
    // Find first valid row with email
    for (var i = 1; i < data.length; i++) {
      var slNo     = data[i][0];  // Column A
      var userName = data[i][1];  // Column B  
      var loginId  = data[i][2];  // Column C
      var password = data[i][3];  // Column D
      var email    = data[i][4];  // Column E
      
      console.log("Testing row " + (i+1) + ":");
      console.log("- Serial: " + slNo);
      console.log("- Name: " + userName);
      console.log("- Login: " + loginId);
      console.log("- Password: " + password);
      console.log("- Email: " + email);
      
      if (email && email.toString().trim() !== "") {
        var emailStr = email.toString().trim();
        var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        
        if (emailRegex.test(emailStr)) {
          console.log("✅ Found valid email: " + emailStr);
          
          // Send test email
          MailApp.sendEmail({
            to: emailStr,
            subject: "🧪 TEST - Your VPN Credentials",
            body: "This is a test email for " + userName + "\n\nLogin: " + loginId + "\nPassword: " + password + "\n\nIf you received this, the system is working!"
          });
          
          console.log("✅ Test email sent successfully!");
          return;
          
        } else {
          console.log("❌ Invalid email format: " + emailStr);
        }
      } else {
        console.log("❌ No email in this row");
      }
    }
    
    console.log("❌ No valid email found in any row");
    
  } catch (error) {
    console.error("❌ Test failed: " + error.toString());
  }
}
