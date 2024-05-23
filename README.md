# From Inquiry to Action: Automating Emails with Excel

Author: Cody

## Background and Goals
This project was created during my work-study job at a talent development team. My role involved:

* Consulting on development opportunities internal & external
* Managed program inquiries following established cross-department process.
* Analyzed inquiry and participation data to improve program development.

**The Challenge: Low Participation**

Despite receiving numerous inquiries daily, our team encountered a concerning trend: a significant gap between those expressing interest and those actually enrolling in programs. Many people inquired, but few signed up for completion. To understand this disconnect and improve program effectiveness, we meticulously tracked both inquiries and program completions.


**Our Goal: Feedback Collection**

As a newly merged team, we wanted to improve the quality of our services and encourage participation in development opportunities. To achieve this, we decided to implement a follow-up email system for program participants.

**Starting Simple with Excel**

Unfortunately, there were no existing software or dedicated systems for sending follow-up emails. So, we opted for a basic solution using Microsoft Excel as a starting point. 

## Solution Design

For the solution part, I also have created a powerpoint for my colleagues for the future use, the readness is also easier, if you would like can skip here, and up for the colorful version. click here.

* **Step 1: Design the Email Template:**

Create a sample email with placeholders for dynamic information, such as [name] and [course name]. These placeholders will be replaced automatically using VBA code. Save the template as a .msg file format.
Example mail:
``` 
Dear [name],

Congratulations on completing the development opportunity for [course name]!

We are confident you gained valuable knowledge and skills during the program. To help us improve future offerings, we'd appreciate your feedback on your experience. 

... (Include a call to action for feedback submission)

Sincerely,

The Talent Development Team
```
* **Step2: Store the Template in a Centralized Location:**

Save the .msg template file in a readily accessible location, such as a shared folder, to ensure easy access for all colleagues.

* **Step 3: Enable the VBA Developer Tab:**

1. Open Microsoft Excel.
2. Go to File > Options.
3. In the Customize Ribbon section on the left, locate the Main Tabs checkbox list.
4. Check the box next to Developer.
5. Click OK.
6. Create and Implement VBA Code:
7. Click on the newly added Developer tab.
8. Click Insert > Module.
9. Paste the VBA code you intend to use for email automation within this module.

_Detailed explanations and the code itself are available in the accompanying document titled "VBA_Code_Explanation.docx"._

_Security: This separate document is not included in this public repository due to security concerns. Please refer to the official Microsoft documentation on VBA security best practices before implementing the code._

* **Step 4: Create and Assign Macro Button**
  
1. Go to Insert > Shapes (choose a shape).
2. Draw the button on your sheet.
3. Right-click the button and select Assign Macro.
4. Choose your VBA code name and click OK.

## Implementation and Results

To sum up, here's how we implemented the Excel-based email follow-up system:

1. **Data Preparation:** We compiled a list of program participants with their names, email addresses, and relevant course details.
2. **Template Integration:** The pre-designed email template (refer to Solution Design) was integrated into the Excel workbook.
3. **VBA Code Implementation:** The VBA code was implemented within an Excel module to automate email sending based on the participant data list. 
4. **Macro Button Creation:** A macro button was created on the Excel sheet for easy execution of the email automation process.

**The Impact of Automated Emails on Talent Development:**

While it's still too early to definitively measure the impact on program enrollment rates due to the recent implementation (a few months), the automated email follow-up system has already demonstrated positive results:

* **Improved Efficiency:**  VBA email systems save time by handling follow-ups within one click, freeing up team members from the tedious task of manually tracking completion dates and copying email content. This automation not only saves time but also adds a personalized touch to our communications.
  
* **Enhanced Program Quality:** Internal feedback mechanisms refine programs based on identified strengths and weaknesses, ensuring continuous improvement. Such as wish of more interactive content and hands-on exerices on certificate courses.
  
* **Greater Pool of Attractive Offerings:**  Incorporating validated external offers recommended by employees enriches our consultancy services, broadening our portfolio. 

* **Comprehensive Feedback Collection:** The automated system goes beyond talent development, enabling efficient feedback gathering during onboarding calls, showcasing its scalability.





