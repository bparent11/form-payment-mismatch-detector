import pandas as pd
import numpy as np
import unicodedata as ud

def normalize_name(name):
    return name.capitalize()

def remove_accents(input_str):
    # Normalisation en forme 'NFKD' pour décomposer les accents
    nfkd_form = ud.normalize('NFKD', input_str)
    # Garder uniquement les caractères non diacritiques (sans accent)
    return ''.join([c for c in nfkd_form if not ud.combining(c)])

cotisantsPSL = pd.read_excel('C:/Users/Lenovo/Documents/Études/Ecole_Ingénieur/BDS/Mandat/algo/form-payment-mismatch-detector/FakeData_Try/data/cotisants.xlsx')
cotisantsPSL = pd.DataFrame({"Nom" : cotisantsPSL['Nom'], "Prénom" : cotisantsPSL['Prénom']})

cotisantsPSL['Nom'] = cotisantsPSL['Nom'].apply(normalize_name)
cotisantsPSL['Prénom'] = cotisantsPSL['Prénom'].apply(normalize_name)

cotisantsPSL['Nom'] = cotisantsPSL['Nom'].apply(remove_accents)
cotisantsPSL['Prénom'] = cotisantsPSL['Prénom'].apply(remove_accents)

cotisantsPSL['Nom Complet'] = cotisantsPSL['Nom'].str.strip() + ' ' + cotisantsPSL['Prénom'].str.strip()
list_cotisantsPSL = cotisantsPSL['Nom Complet'].tolist()

#again ..

form_answer = pd.read_excel('C:/Users/Lenovo/Documents/Études/Ecole_Ingénieur/BDS/Mandat/algo/form-payment-mismatch-detector/FakeData_Try/data/reponsesForm.xlsx')
form_answer = pd.DataFrame({"Nom" : form_answer['Nom'], "Prénom" : form_answer['Prénom']})

form_answer['Nom'] = form_answer['Nom'].apply(normalize_name) #met une majuscule sur la première lettre puis tout en minuscule.
form_answer['Prénom'] = form_answer['Prénom'].apply(normalize_name)

form_answer['Nom'] = form_answer['Nom'].apply(remove_accents) #enlève les accents
form_answer['Prénom'] = form_answer['Prénom'].apply(remove_accents)

form_answer['Nom complet'] = form_answer['Nom'].str.strip() + ' ' + form_answer['Prénom'].str.strip()
list_form = form_answer['Nom complet'].tolist()



list_prompt_payer = []
list_delinquent_payer = []
list_prompt_payer_without_form = []

list_prompt_payer = list(set(list_cotisantsPSL).intersection(set(list_form)))
#we got the prompt payer, now we want to see who answered the form but didn't pay us. Then, we would want to see who paid but didn't fill the form.

#those who answered the form without paying
list_delinquent_payer = list(set(list_form).difference(set(list_cotisantsPSL)))

#those who paid without filling the form
list_prompt_payer_without_form = list(set(list_cotisantsPSL).difference(set(list_form)))

#difference returns the elements that are present in the first object but not in the second.

print("Payeurs qui ont répondu au formulaire :", list_prompt_payer)
print("Répondants qui n'ont pas payé :", list_delinquent_payer)
print("Payeurs qui n'ont pas répondu au formulaire :", list_prompt_payer_without_form)
print("Nombre de payeurs qui ont répondu au formulaire :", len(list_prompt_payer))
print("Nombre de répondants qui n'ont pas payé :", len(list_delinquent_payer))
print("Nombre de payeurs qui n'ont pas répondu au formulaire :", len(list_prompt_payer_without_form))


list_prompt_payer = pd.DataFrame(list_prompt_payer, columns=['a payé et répondu au form'])
list_prompt_payer_without_form = pd.DataFrame(list_prompt_payer_without_form, columns=['a payé sans répondre au form'])
list_delinquent_payer = pd.DataFrame(list_delinquent_payer, columns=['a répondu au form sans payer'])

list_prompt_payer.to_json('data/prompt_payers.json', orient='records')
list_prompt_payer_without_form.to_json('data/prompt_payers_without_form.json', orient='records')
list_delinquent_payer.to_json('data/delinquent_payers.json', orient='records')

#Exporting in xlsx format

max_length = max(len(list_prompt_payer), len(list_prompt_payer_without_form), len(list_delinquent_payer))

list_prompt_payer = list_prompt_payer.reindex(range(max_length))
list_prompt_payer_without_form = list_prompt_payer_without_form.reindex(range(max_length))
list_delinquent_payer = list_delinquent_payer.reindex(range(max_length))

final_df = pd.concat([list_prompt_payer, list_prompt_payer_without_form, list_delinquent_payer], axis=1)

final_df.to_excel("data/final_df.xlsx", index=False)