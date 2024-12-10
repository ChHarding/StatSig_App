# StatSig_App
This guide will walk you through the process of installing, setting up, and using the Statistical Significance Calculator application. This app helps you calculate statistical significance by cycling through stat sig levels to tell you if and at what level a difference is significant, visualize the results, and export them to PowerPoint.

Prerequisites
Before running the app, ensure you have the following installed on your system:
•	Python 3.8+
•	pip (Python package installer)
Required Python Libraries:
The app requires several Python libraries. Run the following command to install all dependencies:
pip install customtkinter scipy matplotlib seaborn pandas python-pptx

________________________________________________________________________________________________________________________________________________ 
Launching the Application
1.	Save the script in a .py file (e.g., significance_calculator.py).
2.	Open a terminal or command prompt.
3.	Navigate to the directory containing the script.
4.	Run the app using the following command:
python significance_calculator.py

 ________________________________________________________________________________________________________________________________________________ 
Main Features Step-by-Step Guide

1.	Input statistical data to calculate significance.

2.	Choose custom bar chart colors.
3.	Visualize results with a dynamically generated graph.
4.	Export analysis to a PowerPoint slide.

________________________________________________________________________________________________________________________________________________  Step-by-Step Process

Step 1: Input Data
1.	Enter Sample Size A and Percentage A in the respective fields.
2.	Enter Sample Size B and Percentage B in their respective fields.
o	Ensure percentages are between 0 and 100.
3.	Optionally, enter a title for your PowerPoint slide if you plan to export.
 
 <img width="200" alt="image" src="https://github.com/user-attachments/assets/a930e1fd-a6ae-4468-8769-eb4e081c4f6f">

 
Step 2: Choose Bar Colors (Optional)
1.	Click Pick Bar Color.
2.	Choose two colors to represent the bars in the chart.

 <img width="203" alt="image" src="https://github.com/user-attachments/assets/b803ec2c-e997-467b-adfb-f386f3ad450b">

 
Step 3: Compute Significance
1.	Click Compute.
o	The app will calculate if the difference between the two groups is statistically significant and at the highest confidence level (80%–99%) it becomes significant at.
o	Results are displayed below the "Compute" button.
 <img width="244" alt="image" src="https://github.com/user-attachments/assets/fb7251db-47a8-4cd9-96a9-de439f600ba8">

Step 4: View and Customize Graph
1.	The graph is displayed at the bottom of the app.
2.	The graph includes:
o	Bars for Percentage A and Percentage B.
o	A significance threshold line, if applicable.
o	A table showing required differences for each significance level.

 <img width="263" alt="image" src="https://github.com/user-attachments/assets/1274eff4-314d-43c9-a7cc-221fee1ae2d5">

Step 5: Export to PowerPoint
1.	Click Export to PowerPoint.
2.	Choose an existing PowerPoint template or create a blank one.
3.	The app will:
o	Add a title (if provided).
o	Include the graph.
o	Add a summary of the results.
4.	Save the PowerPoint file.
5.	
<img width="482" alt="image" src="https://github.com/user-attachments/assets/b6e0e365-0f90-45b5-b3ca-6321b6f6c987">

  ________________________________________________________________________________________________________________________________________________ 
Example Usage
Here’s an example workflow:
1.	Enter:
Sample Size A: 100
Percentage A: 40

Sample Size B: 150
Percentage B: 50

2.	Click Compute:
Output: Significant at 85% level (p-value: 0.0345).
3.	View the graph displaying:
Bars for 40% and 50%.
A red dashed line at the 85% confidence level.
4.	Export results to PowerPoint with the title: Comparison of Marketing Campaigns.

 ________________________________________________________________________________________________________________________________________________ 
5. Troubleshooting
Common Errors:
•	Invalid Input: Ensure all fields are filled and percentages are between 0 and 100.
•	Missing Modules: Install required libraries as mentioned in the Prerequisites section.
•	Export Issues: If exporting fails, ensure the target file isn’t open or locked.

 ________________________________________________________________________________________________________________________________________________ 
6. Known Limitations
1.	Currently supports only two groups for comparison.
2.	Assumes data follows a binomial distribution.
3.	Output graph customization is limited to bar colors.
