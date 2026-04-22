Convert the questions from the attached document into a JSON format compatible with the "survey_kit" library. Note that in step 3. two additional attributes are added "correctAnswer" and "rationale". Add these attributes accordingly from the document as well.

Follow these specific formatting and structural rules:

1. **Split by Level:** Generate separate JSON structures for each care level found in the document (e.g., one JSON for "Primary", one for "Secondary", one for "Tertiary", etc.). If a block contains both Secondary and tertiary, duplicate the questions for each care level.

2. **JSON Structure:** Each JSON must have a root key "steps" containing an array of question objects.

3. **Field Mapping for each question object:**
   - "stepIdentifier": Create a unique ID object (e.g., { "id": "2" }).
   - "type": Set this to "question".
   - "title": Extract the topic header (the content found in the 'green boxes' or section headers, e.g., "Acute rheumatic fever & rheumatic heart disease") and place it here.
   - "text": Place the actual question text here.
   - "correctAnswer": Extract the char specified after the Answer text
   - "rationale": Extract the text after the Rationale: element for each question
   - "answerFormat": An object containing:
     - "type": Set strictly to "single".
     - "textChoices": An array of answer options, each having "text" (the answer description) and "value" (e.g., "A", "B").

4. **Exclusions:** Do not include a "metadata" block.

5. **Example Format:**
   {
     "steps": [
       {
         "stepIdentifier": { "id": "2" },
         "type": "question",
         "title": "Topic Name Here",
         "text": "Actual question text here?",
         "correctAnswer": "Character of the correct answer",
         "rationale": "Rationale of this question",
         "answerFormat": {
           "type": "single",
           "textChoices": [
             { "text": "Option One", "value": "A" },
             { "text": "Option Two", "value": "B" }
           ]
         }
       }
     ]
   }


Always add this block (with some text replacements) as first step:
		{
			"stepIdentifier": {
				"id": "1"
			},
			"type": "intro",
			"title": "<Place name of document here>",
			"text": "Multiple Choice for <Place name of level> Care",
			"buttonText": "Start!"
		},


