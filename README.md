# Microsoft Graph Custom Data Example

## Introduction

This sample shows how to connect a Node.js app to a Microsoft work or school account using the Microsoft Graph API to create/use custom data with calendar events. In addition, the samples uses the Office Frabric UI for styling and formatting the user experience. 

## Prerequisites

To use the this sample, you need the following:

* Node.js version 4 or above.
* Either a Microsoft account or a work or school account

## Register the application

1. Sign into the App Registration Portal using either your personal or work school account.

2. Choose **Add an app**

3. Enter a name for the app, and choose **Create application**.

    The registration page lists the propertiesof your app.

4. Copy the **Application Id**. This is the unique identifier for you app.

5. Under **Application Secrets**, choose **Generate New Password**. Copy the password from the **New password generated** dialog.

    You'll use the application ID and secret to configure the sample app in the next section.

6. Under **Platforms**, choose **Add Platform**.

7. Choose **Web**.

8. Ender [http://localhost:3000/login]() as the Redirect URI.

9. Choose **Save**.

## Build and run the sample

1. Download or clone the Microsoft Graph Custom Data Example

2. Useing you favorite IDE, open **utlis/config.js**.

3. Replace the **clientID** and **clientSecret** placeholder values with the Appplication ID and Application Secret you copied during app registration.

4. In command prompt, run the following command in the root directory. ``npm install``. This installs the project dependencies. 

5. Run the following command to start the development server. ``npm start``.

6. Navigate to ``http://localhost:3000`` in your web browser. 

7. Choose the **Connect** button.

8. Sign in with your personal or work or school account and grant the requested permissions.

## Other information

Make sure to take a look at the **routes/index.js**. You will have to add in event ids to make some of the routes work. Here are the routes to check:
* ``/open``
* ``/open-extensions``
* ``/event-schema-extension``
* ``/get-event-with-outlook-extended-property``
