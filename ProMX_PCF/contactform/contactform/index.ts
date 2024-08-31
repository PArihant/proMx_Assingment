import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class contactform implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _contactId: string | null = null;
    private _notifyOutputChanged: () => void;
    private _jsonData: string = ""; // To store the JSON string
    private _isDirty: boolean = false; // To track if the form has been modified

    // Variables for auto-refresh logic
    private _previousJsonData: string = ""; // To track the last known trial_contactform value
    private _suppressFieldChange: boolean = false; // To suppress field change events during programmatic updates

    constructor() { }

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this._container = container;
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._previousJsonData = context.parameters.trial_contactform.raw || "";

        // Apply CSS for the form
        const style = document.createElement("style");
        style.innerHTML = this._getEmbeddedCss();
        document.head.appendChild(style);

        // Apply HTML for the form
        this._container.innerHTML = this._getEmbeddedHtml();

        // Attach event listeners to form elements
        this._attachEventListeners();

        // Populate the form with contact details if they exist
        this._populateFormIfContactExists();
    }

    // Method to attach event listeners to form elements
    private _attachEventListeners(): void {
        const submitButton = this._container.querySelector("#submitButton") as HTMLButtonElement;
        const cancelButton = this._container.querySelector("#cancelButton") as HTMLButtonElement;

        submitButton.addEventListener("click", this._onSubmit.bind(this));
        cancelButton.addEventListener("click", this._onCancel.bind(this));

        // Add event listeners to form fields to update JSON data on change
        const formFields = this._container.querySelectorAll("input, select");
        formFields.forEach((field) => {
            field.addEventListener("input", () => this._onFieldChange());
        });
    }

    // Event handler for form field changes
    private _onFieldChange(): void {
        if (this._suppressFieldChange) return; // Prevent handling changes triggered programmatically

        // Update JSON data whenever a field changes
        this._updateJsonData();
        this._isDirty = true; // Mark the form as dirty
    }

    // Method to populate the form fields if a contact exists for the account
    private _populateFormIfContactExists(): void {
        const customContext = this._context as CustomContext;
        const accountId = customContext.page.entityId;

        if (accountId) {
            this._context.webAPI.retrieveRecord("account", accountId, "?$select=primarycontactid,_primarycontactid_value,address1_line1,name").then(
                (account: Account) => {
                    const contactReference = account.primarycontactid;

                    if (contactReference && contactReference.contactid) {
                        this._contactId = contactReference.contactid;

                        if (this._contactId) {
                            this._context.webAPI.retrieveRecord("contact", this._contactId, "?$select=firstname,lastname,gendercode,trial_dateofbirth").then(
                                (contact: Contact) => {
                                    // Populate the form with retrieved contact data
                                    this._populateFormFields(contact);
                                },
                                (error) => {
                                    console.error("Error retrieving contact:", error);
                                }
                            );
                        } else {
                            console.error("Contact ID is null or undefined.");
                        }
                    } else {
                        console.error("No primary contact found on the account.");
                    }
                },
                (error) => {
                    console.error("Error retrieving account:", error);
                }
            );
        } else {
            console.error("Account ID is not available.");
        }
    }

    // Method to populate form fields with contact data
    private _populateFormFields(contact: Contact): void {
        const firstNameInput = this._container.querySelector("#firstName") as HTMLInputElement;
        const lastNameInput = this._container.querySelector("#lastName") as HTMLInputElement;
        const genderSelect = this._container.querySelector("#gender") as HTMLSelectElement;
        const dobInput = this._container.querySelector("#dob") as HTMLInputElement;
        const guidInput = this._container.querySelector("#guid") as HTMLInputElement;

        firstNameInput.value = contact.firstname || "";
        lastNameInput.value = contact.lastname || "";
        genderSelect.value = contact.gendercode ? contact.gendercode.toString() : "";
        dobInput.value = contact.trial_dateofbirth ? contact.trial_dateofbirth.substring(0, 10) : "";
        guidInput.value = this._contactId || "";

        // Create initial JSON data
        this._updateJsonData();
    }

    // Method to update JSON data from form fields
    private _updateJsonData(): void {
        const firstNameInput = this._container.querySelector("#firstName") as HTMLInputElement;
        const lastNameInput = this._container.querySelector("#lastName") as HTMLInputElement;
        const genderSelect = this._container.querySelector("#gender") as HTMLSelectElement;
        const dobInput = this._container.querySelector("#dob") as HTMLInputElement;

        // Create JSON data object
        const jsonData: ContactJsonData = {
            first_name: firstNameInput.value,
            last_name: lastNameInput.value,
            gender: genderSelect.value,
            date_of_birth: dobInput.value,
            contact_guid: this._contactId
        };

        // Convert to string and store in _jsonData
        this._jsonData = JSON.stringify(jsonData);

        // Update the bound text field with the JSON data
        this._context.parameters.trial_contactform.raw = this._jsonData;

        this._isDirty = true; // Mark the form as dirty
    }

    // Event handler for the Submit button click
    private _onSubmit(): void {
        this._saveOrUpdateContact();
    }

    // Method to save or update the contact record
    private _saveOrUpdateContact(): void {
        const firstNameInput = this._container.querySelector("#firstName") as HTMLInputElement;
        const lastNameInput = this._container.querySelector("#lastName") as HTMLInputElement;
        const genderSelect = this._container.querySelector("#gender") as HTMLSelectElement;
        const dobInput = this._container.querySelector("#dob") as HTMLInputElement;

        const updatedContact: Partial<Contact> = {
            firstname: firstNameInput.value,
            lastname: lastNameInput.value,
            gendercode: parseInt(genderSelect.value, 10),
            trial_dateofbirth: dobInput.value
        };

        if (this._contactId) {
            this._context.webAPI.updateRecord("contact", this._contactId, updatedContact).then(
                () => {
                    console.log("Contact updated successfully");
                    this._updatePrimaryContactOnAccount(this._contactId!);
                    this._isDirty = false; // Reset dirty flag after saving
                    this._showSuccessNotification(); // Show success notification
                },
                (error) => {
                    console.error("Error updating contact:", error);
                }
            );
        } else {
            this._context.webAPI.createRecord("contact", updatedContact).then(
                (response) => {
                    const newContactId = response.id;
                    console.log("Contact created successfully:", newContactId);
                    this._contactId = newContactId; // Update the contact ID with the newly created one
                    this._updatePrimaryContactOnAccount(newContactId);
                    this._updateJsonData(); // Update the JSON data with the new contact ID
                    this._isDirty = false; // Reset dirty flag after saving
                    this._showSuccessNotification(); // Show success notification
                },
                (error) => {
                    console.error("Error creating new contact:", error);
                }
            );
        }
    }

    // Method to update the primary contact on the associated account
    private _updatePrimaryContactOnAccount(contactId: string): void {
        const customContext = this._context as CustomContext;
        const accountId = customContext.page.entityId;

        if (accountId) {
            const updateAccount: UpdateAccountData = {
                "primarycontactid@odata.bind": `/contacts(${contactId})`
            };

            this._context.webAPI.updateRecord("account", accountId, updateAccount).then(
                () => {
                    console.log("Primary contact updated on account");
                },
                (error) => {
                    console.error("Error updating primary contact on account:", error);
                }
            );
        } else {
            console.error("Account ID is not available.");
        }
    }

    // Method to show success notification after saving
    private _showSuccessNotification(): void {
        const notificationElement = this._container.querySelector("#successNotification") as HTMLDivElement;
        notificationElement.classList.add("show");

        // Hide notification after 3 seconds
        setTimeout(() => {
            notificationElement.classList.remove("show");
        }, 3000);
    }

    // Event handler for the Cancel button click
    private _onCancel(): void {
        this._clearFormFields();

        // Clear the contact lookup field on the Account form
        this._context.parameters.trial_contactform.raw = null;

        // Reset internal state
        this._jsonData = "";
        this._previousJsonData = "";
        this._isDirty = false;
        this._notifyOutputChanged();
    }

    // Method to clear form fields
    private _clearFormFields(): void {
        const firstNameInput = this._container.querySelector("#firstName") as HTMLInputElement;
        const lastNameInput = this._container.querySelector("#lastName") as HTMLInputElement;
        const genderSelect = this._container.querySelector("#gender") as HTMLSelectElement;
        const dobInput = this._container.querySelector("#dob") as HTMLInputElement;
        const guidInput = this._container.querySelector("#guid") as HTMLInputElement;

        this._suppressFieldChange = true; // Suppress field change events during programmatic updates
        firstNameInput.value = "";
        lastNameInput.value = "";
        genderSelect.value = "";
        dobInput.value = "";
        guidInput.value = "";
        this._suppressFieldChange = false;

        this._contactId = null; // Reset internal contactId

        this._jsonData = ""; // Reset _jsonData
    }

    // Method to update form fields from JSON data
    private _updateFormFieldsFromJson(data: ContactJsonData): void {
        const firstNameInput = this._container.querySelector("#firstName") as HTMLInputElement;
        const lastNameInput = this._container.querySelector("#lastName") as HTMLInputElement;
        const genderSelect = this._container.querySelector("#gender") as HTMLSelectElement;
        const dobInput = this._container.querySelector("#dob") as HTMLInputElement;
        const guidInput = this._container.querySelector("#guid") as HTMLInputElement;

        firstNameInput.value = data.first_name || "";
        lastNameInput.value = data.last_name || "";
        genderSelect.value = data.gender || "";
        dobInput.value = data.date_of_birth || "";
        guidInput.value = data.contact_guid || "";

        // Update internal contactId if contact_guid has changed
        if (data.contact_guid && data.contact_guid !== this._contactId) {
            this._contactId = data.contact_guid;
        }

        // Update _jsonData to reflect the new data
        this._jsonData = JSON.stringify(data);
    }

    // Method to update the view when the main form changes
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        const newJsonData = context.parameters.trial_contactform.raw || "";

        if (this._isDirty) {
            this._saveOrUpdateContact(); // Save the contact form when the main form is saved
        }

        if (newJsonData !== this._previousJsonData) {
            if (newJsonData) {
                try {
                    const parsedData: ContactJsonData = JSON.parse(newJsonData);
                    this._suppressFieldChange = true; // Suppress field change events during programmatic updates
                    this._updateFormFieldsFromJson(parsedData);
                    this._suppressFieldChange = false;
                } catch (e) {
                    console.error("Invalid JSON data in trial_contactform:", e);
                }
            } else {
                this._suppressFieldChange = true; // Suppress field change events during programmatic updates
                this._clearFormFields();
                this._suppressFieldChange = false;
            }

            this._previousJsonData = newJsonData; // Update _previousJsonData to the new value
        }
    }

    // Method to return outputs when the main form is saved
    public getOutputs(): IOutputs {
        if (this._isDirty) {
            this._saveOrUpdateContact(); // Save on main form save
        }
        return {
            trial_contactform: this._jsonData
        };
    }

    public destroy(): void {
        // Cleanup logic if needed
    }

    // Method to return the embedded CSS for the form
    private _getEmbeddedCss(): string {
        return `
            .form-container {
                max-width: 600px;
                margin: 0 auto;
                padding: 20px;
                border: 1px solid #e1e1e1;
                border-radius: 4px;
                background-color: #ffffff;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            }
            h3 {
                font-weight: 600;
                margin-bottom: 10px;
            }
            .separator {
                height: 2px;
                background-color: #e1e1e1;
                margin-bottom: 20px;
            }
            .form-row {
                display: flex;
                align-items: center;
                margin-bottom: 15px;
            }
            .form-row label {
                flex: 1;
                font-weight: bold;
            }
            .form-row input, .form-row select {
                flex: 2;
                padding: 8px;
                border: 1px solid #ccc;
                border-radius: 4px;
            }
            .form-actions {
                display: flex;
                justify-content: flex-end;
                margin-top: 20px;
            }
            .btn {
                padding: 10px 20px;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-weight: bold;
            }
            .btn-primary {
                background-color: #0078d4;
                color: white;
                margin-right: 10px;
            }
            .btn:hover {
                opacity: 0.8;
            }
            .footer {
                text-align: center;
                font-size: 12px;
                color: #999;
                margin-top: 20px;
            }
            .notification {
                display: none;
                padding: 10px;
                margin-top: 10px;
                background-color: #d4edda;
                color: #155724;
                border: 1px solid #c3e6cb;
                border-radius: 4px;
            }
            .notification.show {
                display: block;
            }
        `;
    }

    // Method to return the embedded HTML for the form
    private _getEmbeddedHtml(): string {
        return `
            <div class="form-container">
                <h3><span style="vertical-align: middle; margin-right: 10px;">&#128100;</span>Contact Details</h3>
                <div class="separator"></div>
                <form>
                    <div class="form-row">
                        <label for="firstName">First Name</label>
                        <input type="text" id="firstName" name="firstName" />
                    </div>
                    <div class="form-row">
                        <label for="lastName">Last Name</label>
                        <input type="text" id="lastName" name="lastName" />
                    </div>
                    <div class="form-row">
                        <label for="gender">Gender</label>
                        <select id="gender" name="gender">
                            <option value="">--Select--</option>
                            <option value="1">Male</option>
                            <option value="2">Female</option>
                        </select>
                    </div>
                    <div class="form-row">
                        <label for="dob">Date of Birth</label>
                        <input type="date" id="dob" name="dob" />
                    </div>
                    <!-- Hidden Contact GUID Field -->
                    <div class="form-row" style="display: none;">
                        <label for="guid">Contact GUID</label>
                        <input type="text" id="guid" name="guid" disabled />
                    </div>
                    <div class="form-actions">
                        <button type="button" id="submitButton" class="btn btn-primary">Submit</button>
                        <button type="button" id="cancelButton" class="btn">Cancel</button>
                    </div>
                </form>
                <div id="successNotification" class="notification">
                    Contact saved successfully!
                </div>
                <p class="footer">Developed by Arihant Jain</p>
            </div>
        `;
    }
}

// Interface for custom context, including page entityId
interface CustomContext extends ComponentFramework.Context<IInputs> {
    page: {
        entityId: string;
    };
}

// Interface for Account entity with optional primary contact
interface Account {
    primarycontactid?: {
        contactid?: string;
    };
}

// Interface for Contact entity with optional fields
interface Contact {
    contactid?: string;
    firstname?: string;
    lastname?: string;
    gendercode?: number;
    trial_dateofbirth?: string;
}

// Interface for JSON data structure
interface ContactJsonData {
    first_name: string;
    last_name: string;
    gender: string;
    date_of_birth: string;
    contact_guid: string | null;
}

// Interface for updating account's primary contact
interface UpdateAccountData {
    "primarycontactid@odata.bind": string;
}
