import { useEffect, useState } from "react";
import { useMsal } from "@azure/msal-react";
import { InteractionRequiredAuthError, InteractionType, PublicClientApplication } from "@azure/msal-browser";
import { Client } from "@microsoft/microsoft-graph-client";
import { AuthCodeMSALBrowserAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";

// Define a Contact interface for contact information
interface Contact {
  id: string;
  displayName: string;
  emailAddresses: { address: string }[];
  photo?: string; // Optional field for photo URL
}

export const ContactsComponent = () => {
  const { instance, accounts } = useMsal(); // Get MSAL instance and account details
  const [contacts, setContacts] = useState<Contact[]>([]); // State to store fetched contacts
  const [loading, setLoading] = useState(true); // Loading state

  // Ensure we're passing a valid PublicClientApplication instance
  const msalInstance = instance as PublicClientApplication;

  // Initialize Microsoft Graph client with MSAL authentication
  const graphClient = Client.initWithMiddleware({
    authProvider: new AuthCodeMSALBrowserAuthenticationProvider(msalInstance, {
        account: accounts[0],
        scopes: ["Contacts.Read"], // Set required scopes
        interactionType: InteractionType.Redirect, // Include interactionType
      }),
    });

  useEffect(() => {
    const fetchContacts = async () => {
      try {
        // Fetch contacts from the Microsoft Graph API
        const response = await graphClient.api("/me/contacts").get();
        
        // Fetch photos for each contact
        const contactsWithPhotos = await Promise.all(
          response.value.map(async (contact: any) => {
            let photoUrl = "https://via.placeholder.com/150"; // Default photo

            try {
              // Fetch contact's profile picture from Microsoft Graph
              const photoResponse = await graphClient.api(`/me/contacts/${contact.id}/photo/$value`).get();
              const photoBlob = await photoResponse.blob();
              photoUrl = URL.createObjectURL(photoBlob); // Create object URL for photo
            } catch (error) {
              console.log(`No photo available for contact: ${contact.displayName}`);
            }

            return { ...contact, photo: photoUrl }; // Return contact with photo URL
          })
        );

        setContacts(contactsWithPhotos); // Set contacts in state
      } catch (error) {
        // Handle permission/authentication errors
        if (error instanceof InteractionRequiredAuthError) {
          msalInstance.acquireTokenRedirect({ scopes: ["Contacts.Read"] });
        } else {
          console.error("Error fetching contacts:", error);
        }
      } finally {
        setLoading(false); // Stop loading when fetch completes
      }
    };

    fetchContacts();
  }, [msalInstance, graphClient]);

  if (loading) {
    return <p>Loading contacts...</p>; // Show loading message while fetching
  }

  return (
    <div>
      <h1>My Contacts</h1>
      <div style={{ display: "flex", flexWrap: "wrap" }}>
        {contacts.map((contact) => (
          <div
            key={contact.id}
            style={{
              margin: "10px",
              padding: "10px",
              border: "1px solid #ccc",
              borderRadius: "5px",
              textAlign: "center",
            }}
          >
            <img
              src={contact.photo}
              alt={`${contact.displayName}'s profile`}
              style={{
                width: "150px",
                height: "150px",
                borderRadius: "50%",
                objectFit: "cover",
              }}
            />
            <h2>{contact.displayName}</h2>
            {contact.emailAddresses.length > 0 && <p>{contact.emailAddresses[0].address}</p>}
          </div>
        ))}
      </div>
    </div>
  );
};