module.exports = {
    type: "word",
    templateFile: "sample_template.docx", 
    contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    fileExtension: ".docx",
    parseFields: (prettyString) => {
        // Log the incoming string to your terminal so you can see exactly what Jotform sent
        console.log("--- Raw Jotform String Received ---");
        console.log(prettyString);
        console.log("------------------------------------");

        // Updated regex: matches the field label and captures everything up to the end of that line (\r or \n)
       // This tells the engine to stop the moment it hits a comma, a carriage return, or a newline
        const nameMatch        = prettyString.match(/(?:Full\s+)?Name:\s*([^,\r\n]+)/i);
        const emailMatch       = prettyString.match(/Email:\s*([^,\r\n]+)/i);
        const dateMatch        = prettyString.match(/Date:\s*([^,\r\n]+)/i);
        const projectNameMatch = prettyString.match(/Project\s*Name:\s*([^,\r\n]+)/i);
        const projectRefMatch  = prettyString.match(/Project\s*Ref:\s*([^,\r\n]+)/i);

        // Normalize data values safely with clean formatting fallbacks
        const extractedName        = nameMatch        ? nameMatch[1].trim()        : "Client Name";
        const extractedEmail       = emailMatch       ? emailMatch[1].trim()       : "No Email Provided";
        const extractedDate        = dateMatch        ? dateMatch[1].trim().replace(/\s+/g, '-') : "Today";
        const extractedProjectName = projectNameMatch ? projectNameMatch[1].trim() : "Unnamed Project";
        const extractedProjectRef  = projectRefMatch  ? projectRefMatch[1].trim()  : "No Project Reference";

        return {
            projectName: extractedProjectName, 
            submissionDate: extractedDate,
            
            templateData: {
                fullName: extractedName,  
                email: extractedEmail,    
                date: extractedDate,       
                projectName: extractedProjectName, 
                projectRef: extractedProjectRef     
            }
        };
    }
};