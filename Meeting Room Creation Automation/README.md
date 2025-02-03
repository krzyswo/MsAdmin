# Create Exchange Online Room Mailbox  
   
This PowerShell script is used to create a room mailbox in Exchange Online. The room mailbox represents a physical meeting room or resource that can be scheduled for meetings and events.  
   
## Prerequisites  
   
Before running this script, ensure that you have the following:  
   
- Access to Exchange Online  
- Sufficient permissions to create a room mailbox  
- The required PowerShell modules installed  
   
## Usage  
   
1. Connect to Exchange Online by running the following command:  
   ```powershell  
   Connect-ExchangeOnline  
   ```  
   
2. Enter the alias for the mailbox when prompted. For example, `meetingroom-nl-xxx-projectroom`.  
   
3. Enter the domain for the mailbox when prompted. For example, `@example.com`.  
   
4. Enter the name of the room (up to 64 characters) when prompted.  
   
5. Enter the name of the audio device (e.g., Teams) when prompted.  
   
6. Enter the name of the video device (e.g., Teams) when prompted.  
   
7. Enter the name of the display device (e.g., Flatscreen) when prompted.  
   
8. Enter the floor number when prompted.  
   
9. Enter the building number when prompted.  
   
10. Enter the street address when prompted.  
   
11. Enter the room capacity when prompted.  
   
12. Enter the city when prompted.  
   
13. Enter the two-letter country code (e.g., NL) when prompted.  
   
14. Enter the company name when prompted.  
   
15. Enter the department name when prompted.  
   
16. Enter the room number when prompted.  
   
17. Enter the email address of the distribution group when prompted.  
   
18. Choose whether to add yourself as an owner of the distribution group by entering Y or N when prompted.  
   
19. Enter the email address of the distribution group when prompted.  
   
## Parameters  
   
- Alias: The alias for the mailbox.  
- Domain: The domain for the mailbox.  
- Name: The name of the room (up to 64 characters).  
- Audio: The name of the audio device.  
- Video: The name of the video device.  
- Display: The name of the display device.  
- Floor: The floor number.  
- Building: The building number.  
- Street: The street address.  
- Capacity: The room capacity.  
- City: The city.  
- Country: The two-letter country code.  
- Company: The company name.  
- Department: The department name.  
- RoomNumber: The room number.  
- DistributionGroup: The email address of the distribution group.  
- AddOwner: Whether to add yourself as an owner of the distribution group.  
   
## Output  
   
The script will create a new room mailbox in Exchange Online with the specified properties. It will also set calendar processing, place properties, and user properties for the room.  
   
## Notes  
   
- The password for the room mailbox is securely handled.  
- The script waits for 180 seconds before continuing to ensure the mailbox is created.  
- The script checks the place, calendar processing, mailbox, and user properties to verify the settings.  
- Example placeholders for sensitive emails are included in the script. Replace them with the actual email addresses.  
   
## License  
   
This script is released under the [MIT License](LICENSE).
