module.exports = {
    // 1. Core Metadata
    type: "excel",
    templateFile: "manhole_access.xlsx", // File inside 'Templates/Excel Templates/'
    contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    fileExtension: ".xlsx",

    // 2. Data Extractor & Mapper
    parseFields: (prettyString) => {
        console.log("--- Raw Jotform String Received (Excel) ---");
        console.log(prettyString);
        console.log("-------------------------------------------");

        // Helper: Captures full field text through commas until it sees the next "Heading:" pattern
        const extractField = (str, labelRegex) => {
            const pattern = new RegExp(`${labelRegex.source}\\s*([\\s\\S]*?)(?=\\s*,?\\s*[A-Za-z0-9\\s/]+:|$)`, 'i');
            const match = str.match(pattern);
            return match && match[1] ? match[1].trim().replace(/^[\s,]+|[\s,]+$/g, '') : null;
        };

        // Extract fields safely
        const extractedCompany            = extractField(prettyString, /Company:/) || "Client Name";
        const extractedRegion             = extractField(prettyString, /Region:/) || "No Region Provided";
        const extractedDate               = extractField(prettyString, /Date(?:\s*of\s*request\s*submitted)?:/) || "Today";
        const extractedProjectName        = extractField(prettyString, /Project\s*Name:/) || "Unnamed-Project";
        const extractedProjectRef         = extractField(prettyString, /Project\s*Ref:/) || "No-Project-Reference";
        const extractedOtherRef           = extractField(prettyString, /Other\s*refs:/) || "No-Other-Reference";
        const extractedSiteAccessRequired = extractField(prettyString, /Site\s*Access\s*Required\??:/) || "No-Site-Access-Info";
        const extractedSiteName           = extractField(prettyString, /Site\s*name:/) || "No-Site-Name-Info";
        const extractedSiteNameOrAddress  = extractField(prettyString, /Site\s*name\/address\s*where\s*access\s*is\s*required:/) || "No-Site-Name-Or-Address";
        const extractedServiceImpact      = extractField(prettyString, /Service\s*Impact:/) || "No-Service-Impact-Info";
        const extractedChangeType         = extractField(prettyString, /Change\s*Type:/) || "No-Change-Type-Info";
        const extractedReasonForRequest   = extractField(prettyString, /Reason\s*for\s*request:/) || "No-Reason-For-Request-Info"; 
        const extractedScopeOfWork        = extractField(prettyString, /Scope\s*of\s*Work:/) || "No-Scope-Of-Work-Info";
        const extractedRisk               = extractField(prettyString, /Risk(?:\s*Level)?:/) || "No-Risk-Specified";
        const extractedSitesAffected      = extractField(prettyString, /Sites\s*affected:/) || "No-Sites-Affected-Info";
        const extractedCustomerAffected   = extractField(prettyString, /Customers?\s*affected:/) || "No-Customer-Affected-Info";
        const extractedFallbackPlan       = extractField(prettyString, /Fallback\/Rollback\s*plan:/) || "No-Fallback-Plan-Info";

        return {
            projectName: extractedProjectName, 
            submissionDate: extractedDate.replace(/\s+/g, '-'),
            
            // 🎯 Multi-sheet or single-sheet mapping:
            // Adjust "Sheet1" below to match your actual sheet tab name in Excel
            templateData: {
                "Change Application": {
                    "C5":  extractedCompany, 
                    "C7":  extractedDate, 
                    "C9":  extractedRegion, 
                    "B11": extractedSiteAccessRequired, 
                    "C14": extractedChangeType,
                    "C16": extractedServiceImpact,
                    "C18": extractedReasonForRequest, 
                    "C22": extractedScopeOfWork, 
                    "G25": extractedCustomerAffected,
                    "C28": extractedFallbackPlan,
                    "C25": extractedSitesAffected, 
                    "F7":  extractedOtherRef,             
                    "F9":  extractedSiteName,             
                    "F11": extractedSiteNameOrAddress,    
                    "F14": extractedRisk                  
                }
            }
        };
    }
};