# Microsoft Teams Licensing Changes: What You Need to Know    
  
Microsoft introduced significant changes to the licensing structure of Microsoft Teams within the past year, and these changes have had a noticeable impact on how organizations manage and assign Teams licenses, especially for devices used in meeting rooms. As someone working in IT, navigating these updates is essential to ensure smooth operations, compliance, and proper budget planning. Here, I’ll break down what’s changed, who it affects, and how to adjust to the new licensing structure effectively.    
  
---  
   
## Why Was This Change Introduced and When Did It Start?    
  
Microsoft made the decision to introduce a new licensing structure designed to better meet the needs of Teams users and modern room setups. The previous system lacked clarity in terms of features, and many licenses overlapped unnecessarily. By restructuring to include **Basic, Pro, and Shared Device licenses**, Microsoft has made it easier to match licenses to actual use cases, such as small meeting rooms, shared devices, or advanced office spaces that need more robust functionalities.    
  
These changes officially took effect starting **July 1, 2023**, with the goal of supporting hybrid work scenarios while ensuring businesses only pay for the features they truly need. At the same time, it helps reduce inefficiencies caused by legacy licenses that didn’t always align with the complexity of today’s Teams features.    
  
---  
   
## Who Is Affected?    
  
In my experience, these changes directly affect several types of organizations and setups.    
  
- If you’re using **Teams Rooms devices** (like Logitech systems, Crestron hardware, or Surface Hub), then these licensing updates will directly impact how those systems are managed and what features are available.    
- If your organization relies on **Azure Active Directory dynamic groups** or automated systems to assign licenses, these workflows may break due to the new Product IDs introduced with the updated licenses.    
- If you’re transitioning away from legacy licenses, you might face challenges because the new licensing tiers (Basic, Pro, Shared Device) don’t have one-to-one equivalents.    
  
For smaller businesses that rely primarily on standard Teams desktop applications, these changes might be less noticeable unless Teams Rooms devices are a part of their environment.    
  
---  
   
## Consequences for Customers    
  
### Licensing    
  
With the introduction of this new tiered system, the licenses—**Basic, Shared Device, and Pro**—don’t perfectly align with their older counterparts. If you, like many others, were reliant on legacy licenses, you’ll need to take a closer look at your Teams Rooms environment. Devices and features now need to be carefully mapped to the correct new licenses to avoid overpaying for functionality you don’t use, or worse, under-licensing and losing access to critical tools.    
  
### Budget    
  
This new structure can impact budgets, and I’ve seen firsthand how it forces organizations to rethink how they allocate resources:    
  
- **Basic Licenses**: These are fine for small or simpler setups, but there’s a hard cap of **25 per tenant**. Beyond that, you’ll need to switch to Pro licenses, which increases costs. This means Basic might only make sense if you’ve got a handful of smaller rooms that don’t need much functionality.    
- **Shared Device Licenses**: These are designed for shared-use equipment, such as conference telephones or panel displays. However, it’s worth noting these licenses offer limited features and don’t include access to Teams Rooms Pro tools for advanced management.    
- **Pro Licenses**: These are where the full functionality resides. Using these, you get access to features like **Teams Rooms Pro Management Portal**, advanced analytics, and centralized device monitoring. However, these cost noticeably more than legacy licenses, and that additional expense can add up quickly for larger organizations.    
  
### Automation
  
If your organization uses **dynamic groups** to manage license assignments, you’ll need to make manual adjustments. I’ve found that the new licenses have updated **Product IDs**, meaning any automation based on the old IDs will need to be reconfigured to avoid errors.    
  
### Documentation and Cleanup    
  
For those responsible for Teams Rooms deployments, now’s the time to review your records. I recommend creating or updating your inventory of devices in use, along with the features they require. This will help prevent over-provisioning and ensure you’re not paying for Pro licenses where Basic or Shared Device licenses might suffice.    
  
---  
   
## Steps to follow:    
  
Start with an audit of all your meeting room devices. Identify the equipment you’re using in each room—cameras, microphones, touch panels, and so on—and match those details with the features your organization requires. This level of granularity is crucial for determining the most cost-effective license assignments.    

Based on the audit:    
  
- Assign **Shared Device licenses** to shared conference phones or panels where limited functionality is enough.    
- Use **Basic licenses** for smaller and less complex rooms, keeping in mind the **25-license-per-tenant** limit.    
- Apply **Pro licenses** to larger, high-demand rooms that need advanced capabilities like analytics or deeper management options.   
  
If you’re automating license assignments—like using Azure AD dynamic groups—it’s vital to update those workflows. Replace old license Product IDs with the new ones. Test your changes before rolling them out to prevent any disruption to active licenses.    
  
---  
  
### New Product IDs    
  
One challenge I ran into was integrating the new Product IDs into our systems. The old IDs are now irrelevant, so we had to revisit all automation scripts, dynamic groups, and workflows tied to license assignments. It’s not difficult per se, but it’s time-consuming and easy to overlook.    
  
### Limited Functionality in Basic Licenses    
  
- **Basic licenses** are restrictive—they don’t provide access to the **Teams Rooms Pro Management Portal**, and this limits how much visibility and control you have over smaller rooms.    
- Similarly, **Shared Device licenses** may work for telephony or similar equipment, but they severely restrict advanced features.    
  
### Cost   
  
The **Pro licenses** are excellent if you need advanced tools like analytics, monitoring, or premium room support, but the cost can quickly become a burden if you require these for many rooms. Careful budgeting is necessary.    
  
### No One-to-One Mapping    
  
Finally, since the new licenses don’t directly correlate with the old ones, I had to start from scratch when mapping licenses to meeting rooms and devices. This task requires a deep understanding of both your hardware setup and each room’s feature requirements.    
  
---  
   
From my experience, here’s how to approach this transition:    
  
- **Document Everything**: Maintain an accurate inventory of your Teams Rooms setups and align each device with its required functionality. Only assign costly Pro licenses to rooms that truly need premium features.    
- **Evaluate Costs**: Use the migration process to evaluate the financial impact. We worked with our Microsoft representative to ensure we weren’t exceeding budget unnecessarily.    
- **Test Trial Licenses**: Before fully committing to the changes, we tested some licenses using free trials. This allowed us to confirm what worked best for our environment.    
- **Simplify Workflows**: If your license management is automated, ensure your scripts have been updated and tested. I recommend testing your new system in a sandbox environment before applying it to production.    
