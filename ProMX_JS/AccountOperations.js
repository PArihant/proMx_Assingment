// Namespace declaration for PRO_MX
(function (PRO_MX) {}(window.PRO_MX = window.PRO_MX || {}));

// Namespace declaration for AccountOperations within PRO_MX
(function (AccountOperations) {
    /**
     * @name OnPrimaryContactChange
     * @event On primary contact change
     * @description This method populates all PCF Control field values when the primary contact is changed.
     * @param {object} executionContext - The execution context provided by the form event.
     * @return {void}
     */
    AccountOperations.OnPrimaryContactChange = function (executionContext) {
        "use strict";
        
        // Retrieve the form context
        var formContext = executionContext.getFormContext();

        // Get the value of the primary contact lookup field
        var primaryContactLookup = formContext.getAttribute("primarycontactid").getValue();

        if (primaryContactLookup !== null) {
            // Extract the Contact ID and remove curly braces if present
            var contactId = primaryContactLookup[0].id.replace(/[{}]/g, "");

            // Retrieve the contact record using the Web API
            Xrm.WebApi.retrieveRecord("contact", contactId, "?$select=firstname,lastname,gendercode,birthdate").then(
                function (contact) {
                    // Create a JSON object to store the contact details
                    var contactJson = {
                        first_name: contact.firstname || "",
                        last_name: contact.lastname || "",
                        gender: contact.gendercode ? contact.gendercode.toString() : "",
                        date_of_birth: contact.birthdate ? contact.birthdate.substring(0, 10) : "",
                        contact_guid: contactId
                    };

                    // Update the trial_contactform field with the JSON string
                    formContext.getAttribute("trial_contactform").setValue(JSON.stringify(contactJson));

                    // Refresh the trial_contactform control to reflect changes
                    formContext.getControl("trial_contactform").refresh();
                },
                function (error) {
                    // Log any errors that occur during the Web API call
                    console.error("Error retrieving contact:", error.message);
                }
            );
        } else {
            // If no contact is selected, clear the trial_contactform field
            formContext.getAttribute("trial_contactform").setValue(null);

            // Ensure the field gets saved by setting the submit mode to "always"
            formContext.getAttribute("trial_contactform").setSubmitMode("always");
        }
    };
}(window.PRO_MX.AccountOperations = window.PRO_MX.AccountOperations || {}));
