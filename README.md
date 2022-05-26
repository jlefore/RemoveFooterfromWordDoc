# RemoveFooterfromWordDoc
- This was one of my first python projects, and I spent hours researching code, testing, and debugging.
- The genesis for creating this project and why I explored python was as I understood python scripting could change/modify Word documents. 
- I needed to remove my name from the bottom of a Word document footer generated from my financial planning software. In April 2021, the software provider changed the Word document format dramatically. Before this update, I could manually remove my name in two sections at the front of the Word document. The update included multiple section breaks, so I would need to scroll through each page and section and remove my name. This was very time-consuming, sometimes taking up to 10 min.
- The process of using this .py file is straightforward. I upload the Word document from the planning software into a designated folder with the emoneyremovefooter.py file. I run the .py file using the Anaconda Prompt. The code looks to see how many Word documents are in the folder. 
- If only one Word document is in the folder, it replaces my name in the footer(s) with empty space and then opens the Word document (win32). Then I can save it as named back to the folder immediately if I wish. 
- I proofread and format the word document. Then save it and cut it to another folder.
