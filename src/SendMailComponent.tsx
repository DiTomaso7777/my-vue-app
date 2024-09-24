import { useState } from "react";
import { useMsal } from "@azure/msal-react";
import { InteractionRequiredAuthError, InteractionType, PublicClientApplication } from "@azure/msal-browser";
import { Client } from "@microsoft/microsoft-graph-client";
import { AuthCodeMSALBrowserAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";
import './SendMailComponent.css'; // Import the CSS

export const SendMailComponent = () => {
  const { instance, accounts } = useMsal(); // Get MSAL instance and account details
  const [recipient, setRecipient] = useState<string>(""); // State for recipient email
  const [subject, setSubject] = useState<string>(""); // State for email subject
  const [body, setBody] = useState<string>(""); // State for email body
  const [loading, setLoading] = useState(false); // Loading state
  const [success, setSuccess] = useState<boolean | null>(null); // State for success or failure

  // Ensure we're passing a valid PublicClientApplication instance
  const msalInstance = instance as PublicClientApplication;

  // Initialize Microsoft Graph client with MSAL authentication
  const graphClient = Client.initWithMiddleware({
    authProvider: new AuthCodeMSALBrowserAuthenticationProvider(msalInstance, {
      account: accounts[0],
      scopes: ["Mail.Send"], // Set required scope for sending email
      interactionType: InteractionType.Redirect, // Use redirect for interaction
    }),
  });

  const sendMail = async () => {
    setLoading(true); // Start loading
    setSuccess(null); // Reset success state

    const mail = {
      message: {
        subject: subject,
        body: {
          contentType: "Text", // You can also use "HTML" for rich-text emails
          content: body,
        },
        toRecipients: [
          {
            emailAddress: {
              address: recipient,
            },
          },
        ],
      },
      saveToSentItems: true, // Option to save the sent message in "Sent Items"
    };

    try {
      // Send the email using Microsoft Graph API
      await graphClient.api("/me/sendMail").post(mail);
      setSuccess(true); // Set success state on completion
    } catch (error) {
      // Handle permission/authentication errors
      if (error instanceof InteractionRequiredAuthError) {
        msalInstance.acquireTokenRedirect({ scopes: ["Mail.Send"] });
      } else {
        console.error("Error sending email:", error);
        setSuccess(false); // Set failure state
      }
    } finally {
      setLoading(false); // Stop loading when operation completes
    }
  };

  return (
    <div className="send-mail-container">
      <h1>Send New Mail</h1>
      <form
        onSubmit={(e) => {
          e.preventDefault();
          sendMail(); // Trigger sendMail on form submission
        }}
        className="mail-form"
      >
        <div className="form-group">
          <label htmlFor="recipient">Recipient:</label>
          <input
            type="email"
            id="recipient"
            value={recipient}
            onChange={(e) => setRecipient(e.target.value)}
            required
            placeholder="Recipient Email"
          />
        </div>

        <div className="form-group">
          <label htmlFor="subject">Subject:</label>
          <input
            type="text"
            id="subject"
            value={subject}
            onChange={(e) => setSubject(e.target.value)}
            required
            placeholder="Subject"
          />
        </div>

        <div className="form-group">
          <label htmlFor="body">Body:</label>
          <textarea
            id="body"
            value={body}
            onChange={(e) => setBody(e.target.value)}
            required
            placeholder="Type your message here..."
          />
        </div>

        <button type="submit" disabled={loading} className="send-button">
          {loading ? "Sending..." : "Send Mail"}
        </button>
      </form>

      {success === true && <p className="success-message">Email sent successfully!</p>}
      {success === false && <p className="error-message">Failed to send the email. Please try again.</p>}
    </div>
  );
};
