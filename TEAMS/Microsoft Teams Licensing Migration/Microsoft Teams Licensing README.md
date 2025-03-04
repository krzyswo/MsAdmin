# Microsoft Teams Licensing Changes: What You Need to Know

Microsoft has introduced a new licensing structure designed to better align with the needs of Teams users and modern meeting room setups. The previous system lacked clarity, with overlapping licenses that often led to inefficiencies. By restructuring licenses into **Basic, Pro, and Shared Device** tiers, Microsoft has simplified the process of assigning the right license for different use cases, such as small meeting rooms, shared devices, or advanced office spaces requiring more robust functionalities.

These changes officially took effect on **July 1, 2023** ( Global licensing changes for Microsoft 365, Office 365, and Microsoft Teams licensing effective April 1, 2024 ), with the goal of supporting hybrid work environments while ensuring organizations pay only for the features they need. The new structure also reduces inefficiencies caused by legacy licenses that did not always align with the complexity of modern Teams features. 

---

## Who Is Affected?

These changes impact various organizations and setups, particularly those that rely on Teams-integrated devices and automated license management.

- **Teams Rooms Devices**  
  Organizations using **Logitech systems, Crestron hardware, Surface Hub, and similar equipment** will need to adapt to the new licensing model, as it affects system management and feature availability.

- **License Automation**  
  Organizations using **Azure Active Directory (Azure AD) dynamic groups** or automated license assignment workflows may experience disruptions due to the introduction of new Product IDs. Adjustments will be necessary to maintain automation.

- **Legacy License Transitions**  
  The new licensing tiers—**Basic, Pro, and Shared Device**—do not have direct one-to-one equivalents with previous licenses, requiring careful reassessment of licensing needs.

Smaller businesses that primarily use standard Teams desktop applications may notice minimal impact unless they deploy Teams Rooms devices.

---

## Consequences for Customers

### Licensing

The introduction of **Basic, Shared Device, and Pro** licenses requires careful mapping of devices and features to the appropriate tier to avoid overpaying or losing essential functionality.

### Budget Considerations

The new licensing structure affects budgeting and resource allocation:

- **Basic Licenses**  
  Suitable for small or simple setups but limited to **25 per tenant**. Organizations exceeding this limit must upgrade to Pro licenses, increasing costs.

- **Shared Device Licenses**  
  Designed for shared-use equipment, such as conference telephones and panel displays. However, these licenses provide limited functionality and exclude access to **Teams Rooms Pro Management Portal**.

- **Pro Licenses**  
  The most feature-rich option, providing access to **advanced analytics, centralized device monitoring, and the Teams Rooms Pro Management Portal**. However, these licenses are significantly more expensive than legacy options, which can impact larger deployments.

### Automation Adjustments

Organizations using **Azure AD dynamic groups** for automated license assignments must manually update workflows. The **Product IDs** for the new licenses differ from previous versions, requiring updates to automation scripts and configurations.  
**Testing these changes before deployment is essential to avoid disruptions.**

### Documentation and Cleanup

To ensure a smooth transition, organizations should:

- Conduct a **review of all Teams Rooms deployments**.
- Update **device inventories**, including associated features and required licenses.
- Avoid over-provisioning by assigning only the necessary license tier for each setup.

---

## Implementation Guide

### 1. Conduct an Audit

Identify all meeting room devices, including cameras, microphones, and touch panels. Cross-reference this inventory with the necessary features to determine the most cost-effective licensing assignments.

### 2. Assign the Appropriate Licenses

- **Shared Device Licenses**  
  Assign to shared-use equipment like conference phones and panel displays where advanced functionality is unnecessary.

- **Basic Licenses**  
  Use for small, simple meeting rooms, keeping in mind the **25-license-per-tenant limit**.

- **Pro Licenses**  
  Reserve for high-demand rooms that require advanced analytics, monitoring, or management features.

### 3. Update Automation Workflows

For organizations automating license assignments:

- Replace outdated **Product IDs** with the new ones in Azure AD or custom scripts.
- **Test workflow updates** before applying them to production environments to prevent disruptions.

---

## Key Considerations

### New Product IDs

All previous **Product IDs** are now obsolete, requiring updates to all systems that rely on automated license assignments.  
Organizations should verify and adjust their scripts, dynamic groups, and related workflows accordingly.

### Limited Functionality in Basic Licenses

- **Basic licenses** do not provide access to the **Teams Rooms Pro Management Portal**, restricting visibility and control for small meeting rooms.
- **Shared Device licenses** are intended for telephony and similar devices but do not support advanced management features.

### Cost Management

- **Pro licenses** offer comprehensive functionality but come at a higher cost.
- A careful budgeting strategy is necessary to prevent unnecessary expenditures, especially in larger deployments.

### No Direct License Equivalents

Since the new licensing structure does not directly map to the previous system, organizations must reassess their requirements from the ground up.  
This process involves evaluating both the hardware setup and the specific feature needs of each meeting room.

---

## Best Practices and Lesson Learned
- **Document Everything**  
  Maintain an up-to-date inventory of Teams Rooms setups and ensure each device has the correct license.  
  Avoid assigning costly Pro licenses where a Basic or Shared Device license suffices.

- **Evaluate Costs**  
  Review the financial impact of the transition. Engaging with a Microsoft representative can help optimize licensing costs.

- **Test Trial Licenses**  
  Utilize free trial periods to determine the best licensing model before committing to purchases.

- **Simplify Workflows**  
  Ensure automation scripts are updated, tested, and validated in a sandbox environment before full deployment.
