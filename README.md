# customerinsights

The code defines a TranscriptReader class that reads and processes customer service conversation transcripts stored in JSON files. The class has the following key functionality:

read_transcripts(): Reads the JSON transcript files, extracts relevant information (conversation details, sentiment analysis, etc.), and stores them in a list of conversation summaries.
is_valid_conversation(): Validates if a conversation has all the required fields and correct data types.
format_conversation(): Formats the conversation messages into a readable transcript.
summarize_sentiment(): Analyzes the sentiment distribution in the conversation messages.
get_analytics_data(): Aggregates analytics data (issue types, average satisfaction, resolution rate, average duration, sentiment distribution) from the processed conversations.

The code also defines a PresentationFormatter class to apply consistent formatting to slides, and a DataVisualizer class to create various data visualizations (issue trend chart, sentiment donut chart).
The analyze_with_openai() function uses the OpenAI API to generate insights and recommendations based on the customer service conversation data. The insights cover areas like executive summary, product improvement, complaints analysis, marketing insights, top issues, and sentiment analysis.
The PresentationBuilder class is responsible for creating a PowerPoint presentation with the following slides:

Title slide
Executive summary
Slides for each analysis type, with a relevant visualization

The main() function ties everything together by:

Reading the conversation transcripts from the specified directory
Initializing the necessary clients and classes
Generating the insights and recommendations using the OpenAI API
Creating the PowerPoint presentation with the generated content
Saving the presentation to a file

Overall, the code provides a comprehensive solution for analyzing customer service conversation data and presenting the insights in a PowerPoint presentation format.
