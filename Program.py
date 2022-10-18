import os
import numpy as np
import pandas as pd
from pandas import ExcelWriter
from pandas import DataFrame


def reading_csv_files_2nd_stage(main_summery_excel_file: ExcelWriter, excel_file: ExcelWriter,
                                first_file_name_main: str, directory_path: str, df_building: DataFrame):
    measure_err_list: list[str] = list()
    df_build_point = df_building.columns[1]

    df: DataFrame()
    df_measure = DataFrame()
    for file in os.listdir(directory_path):
        if file.endswith(".csv") and not file.startswith("Measure") and not file.startswith(first_file_name_main):
            temp = file.find("--")
            current_file_main_without_suffix = file[:temp]
            df = pd.read_csv(directory_path + "\\" + file)
            if df.columns[1] == df_build_point:
                df_building[current_file_main_without_suffix] = df.iloc[:, 4]
                df_building[current_file_main_without_suffix] = df_building[current_file_main_without_suffix].shift(1)
                df_building.at[0, current_file_main_without_suffix] = df.columns[4]
                df_building[current_file_main_without_suffix] = df_building[current_file_main_without_suffix].shift(9)
                df_building.at[8, current_file_main_without_suffix] = current_file_main_without_suffix
            else:
                measure_err_list.append(current_file_main_without_suffix)
                print("Test : " + file + " point : " + df.columns[1] + " is different then " +
                      first_file_name_main + "( equals to ", df_build_point, ")")
        elif file.endswith(".csv") and file.startswith("Measure"):
            df_measure = pd.read_csv(directory_path + "\\" + file)
    df_building[measure_err_list] = np.nan
    for err in measure_err_list:
        df_building.at[8, err] = err

    new_columns_names = df_building.columns.tolist()[0:3]
    new_columns_names += df_measure.columns.tolist()

    df_building.columns = new_columns_names[:len(df_building.columns)]
    for i in range(0, 3):
        for column_name in new_columns_names[3:len(df_building.columns)]:
            df_building.at[i, column_name] = df_measure.at[i, column_name]

    for column_name in new_columns_names[len(df_building.columns):]:
        df_building[column_name] = df_measure[column_name]
    df_building.insert(3, "", np.nan)

    print()
    print("saving and closing Excels files...")
    print()
    df_building.to_excel(excel_file, sheet_name=first_file_name_main + " summery")
    df_building.to_excel(main_summery_excel_file, sheet_name=os.path.basename(directory_path))


def creating_first_summery_part_by_file(first_file_name: str, signal_df: DataFrame):

    return_df = DataFrame(columns=[signal_df.columns[:3][0], signal_df.columns[:3][1], signal_df.columns[:3][2],
                                   "Time", first_file_name])

    for i in range(3):
        return_df[signal_df.columns[i]] = signal_df.iloc[:, i]

    return_df["Time"] = signal_df.iloc[:, 3]
    return_df["Time"] = return_df["Time"].shift(1)
    return_df.at[0, "Time"] = signal_df.columns[3]
    return_df["Time"] = return_df["Time"].shift(9)
    return_df.at[8, "Time"] = "Time"
    return_df[first_file_name] = signal_df.iloc[:, 4]
    return_df[first_file_name] = return_df[first_file_name].shift(1)
    return_df.at[0, first_file_name] = signal_df.columns[4]
    return_df[first_file_name] = return_df[first_file_name].shift(9)
    return_df.at[8, first_file_name] = first_file_name
    return return_df


def excel_file_creation(path: str):
    writer = pd.ExcelWriter(path+'.xlsx', engine='xlsxwriter')
    return writer


def create_summer_sheet(main_summery_excel_file: ExcelWriter, excel_file: ExcelWriter, directory_path: str,
                        first_file_name: str):
    # first_file_name = "Vout--140k 3--00000"
    signal_df: DataFrame
    for file in os.listdir(directory_path):
        temp = file.find("--")
        clean_foreach_file_name = file[:temp]
        if clean_foreach_file_name == first_file_name and file.endswith(".csv"):
            signal_df = pd.read_csv(directory_path+"\\" + file)
            break
    else:
        raise FileNotFoundError

    df_building = creating_first_summery_part_by_file(first_file_name, signal_df)
    print()
    reading_csv_files_2nd_stage(main_summery_excel_file, excel_file, first_file_name, directory_path, df_building)


def create_folder_sheets(excel_file: ExcelWriter, directory_path: str):
    df: DataFrame()
    for file in os.listdir(directory_path):
        if file.endswith(".csv") and not file.startswith("Measure"):
            df = pd.read_csv(directory_path + "\\" + file)
            temp = file.find("--")
            file = file[:temp]
            df.to_excel(excel_file, sheet_name=file)
            print(file)
    print()


def searching_filed_by_summery(main_folders_path: str):
    first_file_name = input("\nEnter a name file (without .csv extensionS) for the first column: " + "\n" +
                            "(no Extension only bare name file like : Vout)\n" +
                            "Note that the main folder should contain only sub folders that contains excel files\n")
    print()
    flag = True
    first_folders_name = str()
    for file in os.listdir(main_folders_path):
        if not file.__contains__(".") and not file == "exe":
            first_folders_name = "\\" + file
            break
    while flag:
        try:
            for file in os.listdir(main_folders_path + first_folders_name):
                temp = file.find("--")
                clean_foreach_file_name = file[:temp]
                if clean_foreach_file_name == first_file_name and file.endswith(".csv"):
                    flag = False
                    break
            else:
                raise FileNotFoundError
        except FileNotFoundError as e:
            print(e)
            print("Coundn't find the file specified try entering the name again -  ")
            first_file_name = input("Enter the name of the file (without .csv extension) for the first column\n")
    return first_file_name


def iterating_all_sub_folder_in_main_folder():
    main_folders_path = os.getcwd()
    dirlist = os.listdir(main_folders_path)
    main_summery_excel_file: ExcelWriter = excel_file_creation(main_folders_path + "\\" +
                                                               os.path.basename(main_folders_path))
    clean_first_file_name = searching_filed_by_summery(main_folders_path)

    for directory in dirlist:
        if not directory.__contains__(".") and directory != "exe":
            excel_file: ExcelWriter = excel_file_creation(main_folders_path + "\\" + directory + "\\" + directory)
            print("Creating folder's excel files sheets... ")
            create_folder_sheets(excel_file, main_folders_path + "\\" + directory)
            print("Creating Summery Sheets...")
            create_summer_sheet(main_summery_excel_file, excel_file, main_folders_path + "\\" + directory,
                                clean_first_file_name)
            excel_file.close()
    main_summery_excel_file.close()


if __name__ == "__main__":
    iterating_all_sub_folder_in_main_folder()
    input("Enter any key and enter to exit the Program")
# ===============================================================================================================


# def create_folders_excels_file_name(main_folders_name: str):
#     summery_file_name = "all_"
#     temp_summery_file_name_a = main_folders_name.find("_")
#     temp_summery_file_name_b = main_folders_name.find("m")
#     summery_file_name += main_folders_name[temp_summery_file_name_a+1:temp_summery_file_name_b]
#     main_folders_name = main_folders_name[temp_summery_file_name_b:]
#     temp_summery_file_name_a = main_folders_name.find("=")
#     summery_file_name += "µH"
#     main_folders_name = main_folders_name[temp_summery_file_name_a:]
#     summery_file_name += "_Vin"
#     temp_summery_file_name_a = main_folders_name.find("=")
#     temp_summery_file_name_b = main_folders_name.find("_")
#     summery_file_name += main_folders_name[temp_summery_file_name_a+1:temp_summery_file_name_b]
#     main_folders_name = main_folders_name[temp_summery_file_name_b:]
#     temp_summery_file_name_a = main_folders_name.find("=")
#     main_folders_name = main_folders_name[temp_summery_file_name_a:]
#     summery_file_name += "_Vout"
#     summery_file_name += main_folders_name[1:]
#     return summery_file_name
