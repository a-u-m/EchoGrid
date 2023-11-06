import os
from flask import Flask, request, jsonify
from flask_cors import CORS
import spacy

app = Flask(__name__)
CORS(app)

model_path = os.path.join(os.path.dirname(__file__), "../model/model-best/model-best")
nlp_ner = spacy.load(model_path)

def extract_entities(text):
    values = []
    cells = []
    column = []
    row = []
    type_labels = []

    doc = nlp_ner(text)

    for ent in doc.ents:
        if ent.label_ == "VALUE":
            values.append(ent.text)
        elif ent.label_ in ("INSERT", "DELETE", "REPLACE", "MERGE", "BOLD", "ITALIC"):
            type_labels.append(ent.label_)
        elif ent.label_ == "CELL":
            cells.append(ent.text)
        elif ent.label_ == "COLUMN":
            column.append(ent.text)
        elif ent.label_ == "ROW":
            row.append(ent.text)

    return {
        'values': values,
        'type_labels': type_labels,
        'cells': cells,
        'column': column,
        'row': row
    }

@app.route('/predict', methods=['POST'])
def predict():
    try:
        data = request.get_json()
        text = data['text']
        print(text)
        extracted_entities = extract_entities(text)
        print(extract_entities)
        return jsonify(extracted_entities)

    except Exception as e:
        return jsonify({'error': str(e)})

if __name__ == '__main__':
    app.run(debug=True)
