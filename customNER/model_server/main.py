import os
from flask import Flask, request, jsonify
from flask_cors import CORS
import spacy
import spacy_transformers
app = Flask(__name__)
CORS(app)

model_path = os.path.join(os.path.dirname(__file__), "../model/best_Model")
nlp_ner = spacy.load(model_path)

def extract_entities(text):
    values = []
    cells = []
    column = []
    row = []
    chart = []
    type_labels = []

    doc = nlp_ner(text)

    for ent in doc.ents:
        if ent.label_ == "VALUE":
            values.append(ent.text)
        elif ent.label_ in ("INSERT", "DELETE", "REPLACE", "MERGE", "BOLD", "ITALIC","SUMALL","MULTIPLYALL"):
            type_labels.append(ent.label_)
        elif ent.label_ == "CELL":
            cells.append(ent.text)
        elif ent.label_ == "COLUMN":
            column.append(ent.text)
        elif ent.label_ == "ROW":
            row.append(ent.text)
        elif ent.label_ == "CHART":
            type_labels.append(ent.label_)
            chart.append(ent.text)

    return {
        'values': values,
        'type_labels': type_labels,
        'cells': cells,
        'column': column,
        'row': row,
        'chart': chart
    }

@app.route('/predict', methods=['POST'])
def predict():
    try:
        data = request.get_json()
        text = data['text']
        print(text)

        nlp = spacy.load("en_core_web_sm")

        number_mapping = {
            'zero': '0',
            'one': '1',
            'two': '2',
            'three': '3',
            'four': '4',
            'five': '5',
            'six': '6',
            'seven': '7',
            'eight': '8',
            'nine': '9'
        }

        # Regular expression to match text representations of numbers
        import re
        pattern = re.compile(r'\b(' + '|'.join(number_mapping.keys()) + r')\b', re.IGNORECASE)

        # Replace text representations with numerical values
        text = pattern.sub(lambda x: number_mapping[x.group().lower()], text)

        #input_text.lower()
        # Tokenize and process the input text
        doc = nlp(text)

        # Initialize a list to store sentences
        sentences = []

        # Keywords to split sentences
        split_keywords = ["and", "or", "then", "but", ",", "Now", "now"]

        current_sentence = ""

        for token in doc:
            if token.text in split_keywords:
                if current_sentence:
                    sentences.append(current_sentence.strip())
                current_sentence = ""
            else:
                current_sentence += " " + token.text

        # Add the last sentence
        if current_sentence:
            sentences.append(current_sentence.strip())

        # Now 'sentences' is a list of separated sentences that you can use with your model
        # Print the separated sentences
        #for i, sentence in enumerate(sentences, start=1):
            #print(f"Sentence {i}: {sentence}")

        print(sentences)


        json_list = []
        
        for sentence in sentences:
            json_list.append(extract_entities(sentence))
        
        print(json_list)
        import json
        # json_string = json.dumps(json_list, indent=2)
        # print(json_string)
        
        # extracted_entities = extract_entities(text)
        # print(extract_entities)
        return jsonify(json_list)

    except Exception as e:
        return jsonify({'error': str(e)})

if __name__ == '__main__':
    app.run(debug=True)
