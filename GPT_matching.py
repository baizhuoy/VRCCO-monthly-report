from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.svm import SVC
from sklearn.metrics.pairwise import cosine_similarity
import pandas as pd


# Input master data list as training data
mst = pd.read_excel('//VRCCO-24/Users/Michael/Desktop/INV EXCEL FILES/Monthly Inv Report.xlsx', sheet_name=-1)
name = mst.iloc[:, :3]
s = name['Supplier'].unique()

# Organize the data by supplier
supplier_data = {}
for i in s:
    pn = list(name[name['Supplier'] == i]['Product'])
    sn = list(name[name['Supplier'] == i]['Supplier Name'])
    if len(pn)==1:
        continue
    else:
        supplier_data[i] = (pn, sn)

# Train a model for each supplier
trained_models = {}
trained_vectorizers = {}

for supplier, (product_names, supplier_names) in supplier_data.items():
    vectorizer = TfidfVectorizer()
    tfidf_product_names = vectorizer.fit_transform(product_names)
    svm = SVC(kernel='linear')
    svm.fit(tfidf_product_names, supplier_names)
    trained_models[supplier] = svm
    trained_vectorizers[supplier] = vectorizer

# Read the new Excel file with many supplier names and product names
new_data = pd.read_excel("C:\Inv Data\Purchase History\Purchase History.xlsx",sheet_name=0)

# # Organize the new data by supplier
new_supplier_data = {}

new_supplier_data['MWI'] =list(new_data['Description'])
# Threshold for similarity

# Keep track of the matches for each product name in the new Excel file
matches = []

# Iterate through the suppliers and product names in the new Excel file
for product_name in list(name[name['Supplier'] == 'MWI']['Product']):
    best_similarity = -1
    best_match_name = None

    for supplier, new_supplier_names in new_supplier_data.items():
        for new_supplier_name in new_supplier_names:

            # Use the previously trained models to predict the supplier name
            vectorizer = trained_vectorizers[supplier]
            tfidf_new_product_name = vectorizer.transform([product_name])
            predicted_supplier_name = trained_models[supplier].predict(tfidf_new_product_name)



            # Calculate the similarity between the predicted supplier name and the new product name
            similarity = cosine_similarity(vectorizer.transform([new_supplier_name]),
                                           vectorizer.transform(predicted_supplier_name))



            # Keep track of the best match
            if similarity > best_similarity:
                best_similarity = similarity
                best_match_name = new_supplier_name

        # Add the match (or blank if no match) to the results
        matches.append((product_name, best_match_name, predicted_supplier_name,best_similarity))


df = pd.DataFrame(matches, columns=['Product Name', 'Supplier Name','Predicted','similarity'])

# 将DataFrame导出到Excel文件
output_file_path = 'C:\AI\AI.xlsx'
df.to_excel(output_file_path, index=False)