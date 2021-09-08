import pandas as pd


def read_spreadsheet_and_return_species_and_mirna(path):
    reader = pd.read_excel(path)

    species_list = reader["SPECIES"].tolist()
    mirna_id_list = reader["MIRNA ID"].tolist()

    SPECIES_INTEREST_LIST = ["Homo sapiens",
                             "Mus musculus", "Rattus norvegicus"]

    interest_list = []

    for index, species in enumerate(species_list):
        if species in SPECIES_INTEREST_LIST:
            interest_list.append(
                {
                    'species': species,
                    'mirna': mirna_id_list[index]
                }
            )

    return interest_list
