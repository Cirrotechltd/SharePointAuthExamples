# SharePoint API Console Application

This is a C# console application that interacts with SharePoint Online using the SharePoint REST API. The application uses an Azure AD App Registration with application permissions and certificate-based authentication to perform various operations like retrieving site details, creating lists, and fetching list items.

## Features

- Retrieve SharePoint site information.
- Execute custom SharePoint REST API calls.
- Create new SharePoint lists.
- Retrieve all items from a specific SharePoint list.

## Prerequisites

### Azure App Registration:

- Create an app registration in Azure Active Directory.
- Assign Application Permissions for SharePoint:
  - `Sites.FullControl.All` or `Sites.Manage.All`.
- Upload a certificate (used for authentication).
- Grant admin consent to the app.

### Certificate Setup:

- Generate a certificate using a tool like OpenSSL or PowerShell.
- Install the certificate in the Windows Certificate Store under **CurrentUser → Personal**.
- Note the certificate thumbprint for use in the application.

### Environment:

- .NET 6 or later installed.
- Access to a SharePoint Online site.

### Installation:

1. Clone the repository  
2. Install dependencies: Ensure you have the required .NET libraries, such as Microsoft.Identity.Client  
3. Open the project in Visual Studio or your preferred IDE  
4. Update Program.cs with your Azure App Registration details:  
   - tenantId: Your Azure AD tenant ID  
   - clientId: Your app registration's client ID  
   - certificateThumbprint: The thumbprint of your uploaded certificate  
   - sharepointSiteUrl: The URL of your SharePoint site  
5. Build and run the application:  
   - `dotnet build`  
   - `dotnet run`

### Cert creation

### Generate a private key
\```bash
openssl genrsa -out privatekey.pem 2048
\```

### Create a Certificate Signing Request (CSR)
\```bash
openssl req -new -key privatekey.pem -out certrequest.csr
\```

### Create a self-signed certificate
\```bash
openssl x509 -req -days 730 -in certrequest.csr -signkey privatekey.pem -out certificate.crt
\```

### Combine certificate and private key into a PFX file
\```bash
openssl pkcs12 -export -out certificate.pfx -inkey privatekey.pem -in certificate.crt
\```




