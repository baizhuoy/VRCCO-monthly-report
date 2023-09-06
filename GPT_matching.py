import pandas as pd
from gensim.models import FastText
from fractions import Fraction
from sklearn.svm import SVC
from sklearn.model_selection import RandomizedSearchCV
from sklearn.metrics import classification_report
from sklearn.model_selection import train_test_split
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.model_selection import ShuffleSplit
import numpy as np
import re

# Set for several suppliers 
# def get_most_similar_product(supplier, new_product_name, new_supplier_names, best_similarity, best_match_name):
#     fasttext_model = trained_models[supplier]
#     svm_model = trained_svms[supplier]
#
#     # Transform new product name using FastText model
#     transformed_name = fasttext_model.wv[new_product_name]
#
#     # Predict using the trained SVC model
#     predicted_supplier_name = svm_model.predict([transformed_name])[0]
#
#     # Calculate the similarity between the predicted supplier name and the new supplier name
#     for supplier_name in new_supplier_names:
#         similarity = cosine_similarity(fasttext_model.wv[preprocess_text(supplier_name)].reshape(1, -1),fasttext_model.wv[predicted_supplier_name].reshape(1, -1))
#         if similarity > best_similarity and similarity>0.5:
#             best_similarity = similarity
#             best_match_name = supplier_name
#         print(new_product_name,similarity,supplier_name,best_match_name,predicted_supplier_name)
#     return best_similarity, best_match_name

def preprocess_text(text):
    text_with_numbers = text

    # Convert fractions in the form of 'a/b' to decimals
    fractions = re.findall(r'\b\d+/\d+\b', text_with_numbers)
    for fraction in fractions:
        numerator, denominator = map(int, fraction.split('/'))
        if denominator != 0:
            decimal_value = str(Fraction(fraction))
            text_with_numbers = text_with_numbers.replace(fraction, decimal_value)
    # Replace '' with inch
    text_with_numbers = re.sub(r'"', 'in', text_with_numbers)
    # Remove punctuation except percentage sign and dot
    text_with_numbers = re.sub(r'[^\w\s\.%]', '', text_with_numbers)
    # Replace digits with a special token to emphasize numbers
    text_with_numbers = re.sub(r'(\b\d+\b)', r'DIGIT_\1_DIGIT', text)
    return text_with_numbers


# Input master data list as training data
mst = pd.read_excel('//Monthly Inv Report.xlsx',sheet_name=-1)
name = mst.iloc[:, :3]
s = name['Supplier'].unique()

# Organize the data by supplier
supplier_data = {}
for i in s:
    pn = list(name[name['Supplier'] == i]['Product'])
    sn = list(name[name['Supplier'] == i]['Supplier Name'])
    if len(pn) < 3:
        continue
    else:
        supplier_data[i] = (pn, sn)

# Train a FastText model for each supplier
trained_models = {}
trained_svms = {}

for supplier, (product_names, supplier_names) in supplier_data.items():
    # Train a FastText model for the entire dataset
    sentences = [name for name in product_names]
    model = FastText(sentences, vector_size=100, window=5, min_count=1, workers=4)
    trained_models[supplier] = model

    # Convert product names to vectors
    x = [model.wv[name] for name in sentences]
    y = [preprocess_text(supplier_name) for supplier_name in supplier_names]

    # Split data into training and testing
    # X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

    # Define the hyperparameters and their possible values
    param_dist = {
        'C': np.logspace(-3, 2, 6),
        'gamma': ['scale', 'auto'] + list(np.logspace(-2, 2, 5)),
        'kernel': ['linear', 'rbf', 'poly', 'sigmoid']
    }
    # Define the SVC model
    svc = SVC()

    # Initialize RandomizedSearchCV
    try:
        search = RandomizedSearchCV(svc, param_distributions=param_dist, n_iter=50, cv=5, verbose=1, n_jobs=-1, random_state=42,error_score='raise')
        search.fit(x,y)
        trained_svms[supplier] = search.best_estimator_
    except ValueError:
        search = RandomizedSearchCV(svc, param_distributions=param_dist, n_iter=50, cv=ShuffleSplit(n_splits=1, test_size=0.25), verbose=1, n_jobs=-1,random_state=42,error_score='raise')
        search.fit(x, y)
        trained_svms[supplier] = search.best_estimator_




# Read the new Excel file with many supplier names and product names
new_data = pd.read_excel("C:\Purchase History.xlsx",sheet_name=0)

# Organize the new data by supplier
new_supplier_data = {}

new_supplier_data['MWI'] =list(new_data['Description'])

matches = []
# Iterate through the suppliers and product names in the new Excel file
for product_name in list(name[name['Supplier'] == 'MWI']['Product']):
    best_similarity = -1
    best_match_name = None

  #  best_similarity, best_match_name = get_most_similar_product('MWI', product_name, new_supplier_data['MWI'],best_similarity, best_match_name)
    supplier='MWI'
    fasttext_model = trained_models[supplier]
    svm_model = trained_svms[supplier]

    # Transform new product name using FastText model
    transformed_name = fasttext_model.wv[product_name]

    # Predict using the trained SVC model
    predicted_supplier_name = svm_model.predict([transformed_name])[0]

    # Calculate the similarity between the predicted supplier name and the new supplier name
    for supplier_name in new_supplier_data['MWI']:
        adj_supplier_name=preprocess_text(supplier_name)
        similarity =np.clip(cosine_similarity(fasttext_model.wv[adj_supplier_name].reshape(1, -1),fasttext_model.wv[predicted_supplier_name].reshape(1, -1)),-1,1)
        if similarity > best_similarity and similarity[0] > 0.5:
            best_similarity = similarity
            best_match_name = supplier_name
            print(product_name, similarity, supplier_name, best_match_name, predicted_supplier_name)

    # Add the match (or blank if no match) to the results
    matches.append((product_name, best_match_name, best_similarity))

df = pd.DataFrame(matches, columns=['Product Name', 'Supplier Name', 'similarity'])

# Output matched data
output_file_path = 'C:\AI\AI.xlsx'
df.to_excel(output_file_path, index=False)
