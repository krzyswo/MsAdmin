# How to Create Your Personal Assistant

If you struggle with presenting your skills and experience during interviews, a personal interview assistant can be a valuable tool. This assistant will introduce what you do and highlight your expertise when needed. You can host it on Azure and deploy it using Microsoft's AI solutions. Here’s how you can set it up.

## Prerequisites

Before starting, ensure you have:

- An active Microsoft Azure subscription
- A GitHub account (which will store the assistant’s reference data)

## Deploying Your Assistant

Microsoft provides a [Chat with Your Data solution accelerator](https://github.com/Azure-Samples/chat-with-your-data-solution-accelerator?tab=readme-ov-file) that allows you to create an AI-powered assistant. You can deploy it in two ways: a quick, template-based method or a more advanced method using Visual Studio Code (which will be covered in future articles). Today, we will focus on the easy and fast deployment approach.

### 1. Quick Deployment Using Azure

The fastest way to deploy your assistant is by using the **Deploy to Azure** button, which will automatically set up the necessary resources.

**Steps:**

1. Click the **Deploy to Azure** button on the [Chat with Your Data solution accelerator page](https://github.com/Azure-Samples/chat-with-your-data-solution-accelerator?tab=readme-ov-file).
2. Follow the guided steps to provision the required Azure resources.
3. Configure the application with your preferred settings (e.g., supported languages).

**Mandatory Input:**

- **Setting Location in Deployment**: Select a location such as **East US 2** for your deployment.
- **Setting Name in Deployment**: Choose a name for your assistant, like **interviewassistant**, with a maximum of 20 characters.

When you are deploying resources like your AI assistant, the Location setting will typically be configured in the deployment template, such as:

> **Direct Deployment Template**: You can directly use this [Azure deployment template](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2FAzure-Samples%2Fchat-with-your-data-solution-accelerator%2Frefs%2Fheads%2Fmain%2Finfra%2Fmain.json) to begin the deployment process.

### 2. Integrating Your GitHub Repositories

Once your application is deployed, it needs access to your data. The assistant can reference your GitHub repositories to understand your projects and skills.

![image](https://github.com/user-attachments/assets/f2385eb5-6dd1-438e-a9d8-b00634a0ad7d)


**Steps:**

1. Open your assistant in a browser.
2. You can access the repository's file tree by making a GET request to the following API endpoint:

   This code example will fetch your public repositories from GitHub and print out the list of URLs:  
   `https://api.github.com/repos/{owner}/{repo}/git/trees/{branch}?recursive=1`

   In my case, it would be for example:  
   `https://api.github.com/repos/krzyswo/MsAdmin/git/trees/main?recursive=1`

   ![image](https://github.com/user-attachments/assets/d9a7c8ce-e91e-4431-a613-0ae212aa632e)


4. Upload the list of repository URLs to the admin tool in the **Chat with Your Data** solution accelerator.
   ![image](https://github.com/user-attachments/assets/f9f213cc-b105-429e-8e39-aed6aee2237e)

5. The system will process and index the content of the repositories.

### 3. Verifying Data Processing

To ensure everything has been properly uploaded and indexed:

1. Go to the **Explorer** page in your assistant and check that all repositories have been fully processed.
2. Navigate to the **User Chat** section and interact with the assistant.
3. Ask questions related to your projects and skills to confirm the assistant retrieves the correct information.
![image](https://github.com/user-attachments/assets/ed96a8f2-aace-48b8-9e78-61da621a63ef)
![image](https://github.com/user-attachments/assets/77661533-19f9-4b14-82a6-971146df17a7)


## Conclusion

With this setup, you now have a personal interview assistant that can provide insights into your work and skills whenever needed. As you continue developing it, consider integrating more advanced AI features or refining the way it presents your experience.

Stay tuned for future articles where we will explore deploying the assistant using Visual Studio Code and additional customization options!
