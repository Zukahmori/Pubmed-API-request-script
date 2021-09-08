import requests
import xmltodict
from read_spreadsheet import read_spreadsheet_and_return_species_and_mirna
import pandas as pd

interest_list = read_spreadsheet_and_return_species_and_mirna(
    "planilha-miRNA.ods")


count_file_list = []
miRNA_file_list = []
species_file_list = []
id_file_list = []
interest_list_length = len(interest_list)

for index, item in enumerate(interest_list):
    print(f'Item atual - {index}\nTotal - {interest_list_length}')
    miRNA = item['mirna']
    species = item['species'].split(' ')
    BASE_URL = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi?term={miRNA}%5BAll%20Fields%5D%2Bepilepsy%5BAll%20Fields%5D%2B{species[0]}%2B{species[1]}%5BAll%20Fields%5D"

    response = requests.get(BASE_URL)

    if response.status_code != 200:
        raise Exception('Requisição falhou')

    dictionary = xmltodict.parse(response.content)

    result_count = int(dictionary['eSearchResult']['Count'])

    if result_count > 0:
        id_list = dictionary['eSearchResult']['IdList']['Id']

        if result_count >= 2:
            id_list = ','.join(id_list)

        count_file_list.append(result_count)
        miRNA_file_list.append(miRNA)
        species_file_list.append(" ".join(species))
        id_file_list.append(id_list)

    else:
        count_file_list.append(result_count)
        miRNA_file_list.append(miRNA)
        species_file_list.append(" ".join(species))
        id_file_list.append('NÂO ENCONTRADO')

    write_dictionary = {'Count': count_file_list, 'miRNA': miRNA_file_list,
                        'Species': species_file_list, 'Article IDs': id_file_list}

    df = pd.DataFrame(data=write_dictionary)

    df.to_excel('microRNA-Results.ods')
