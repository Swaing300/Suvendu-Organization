document.getElementById('loiForm').addEventListener('submit', function(event) {
    event.preventDefault();
  
    // Get form values
    const senderName = document.getElementById('senderName').value;
    const senderEmail = document.getElementById('senderEmail').value;
    const companyName = document.getElementById('companyName').value;
    const subject = document.getElementById('subject').value;
    const position = document.getElementById('position').value;
    const phone = document.getElementById('phone').value;
    const address = document.getElementById('address').value;
    const loiContent = document.getElementById('loiContent').value;
  
    // Get the signature image file
    const signatureFile = document.getElementById('signature').files[0];
  
    const reader = new FileReader();
  
    reader.onload = function(event) {
      const signatureDataURL = event.target.result;
  
      // Generate the content for the Word document including the signature image
      const content = `
        <h1 style="text-align: center;">Letter of Intent</h1>
        <p><strong>Date:</strong> ${new Date().toDateString()}</p>
        <p><strong>Recipient:</strong> Suvendu Organization</p>
        
        <p><strong>Dear Suvendu Organization,</strong></p>
        
        <p>I am writing this letter to express my intent regarding <strong>${subject}</strong>.</p>
        
        <p>${loiContent}</p>
        
        <p><strong>Sincerely,</strong></p>
        <p>${senderName}</p>
        <p>(${senderEmail})</p>
        
        <p><img src="${signatureDataURL}" alt="Signature" style="width: 200px; height: auto;"></p>
        
        <p><strong>${companyName}</strong></p>
        <p><strong>${position}</strong></p>
        <p><strong>Phone:</strong> ${phone}</p>
        <p><strong>Address:</strong> ${address}</p>
      `;
  
      // Convert HTML content to Word document
      const docx = window.htmlDocx.asBlob(content);
      saveAs(docx, 'Letter_of_Intent.docx');
    };
  
    // Read the signature image file as data URL
    if (signatureFile) {
      reader.readAsDataURL(signatureFile);
    }
  });
  