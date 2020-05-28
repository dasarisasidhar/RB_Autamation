import pandas as pd
from datetime import datetime

def generate(data):

    """
      generates file takes input dataframe should contain Primary Used by column
      return/saves .xlsx file by coonverting it to anaffe charging formate

    """

    print("Total licenses details given: ", len(data))

    print("Total number of unique user licenses: ", len(data["Primary Used by"].unique()))

    #get name

    name =  input("Enter name of the Product: ")

    #get lasel

    lasel =  input("Enter lasel number: ")

    duplicated_details = data[data.duplicated(['Primary Used by'])]  

    duplicated_details.to_excel("duplicate_user_details.xlsx")

    # remove null values
    data = data.dropna(how='any', axis=0)

    # remove duplicaes
    data.drop_duplicates(['Primary Used by'], inplace= True)


    print("Total licenses after removing duplicats and null values: ",len(data))


    temp_data = data.groupby(["RecCC"]).size().reset_index(name='counts')

    print("Total used licenses: ",temp_data.counts.sum())

    RecCC1 = list()
    RecCC2 = list()

    for i in temp_data["RecCC"]:
        i = str(i)
        RecCC1.append(i[:3])
        RecCC2.append(i[3:])

    temp_data["RecCC1"] = RecCC1
    temp_data["RecCC2"] = RecCC2

    counts = list()

    for i in temp_data["counts"]:
        counts.append(str(i)+".00000")

    temp_data["counts"] = counts

    if datetime.now().month < 10:
        month_year = str(datetime.now().year)+"0"+str(datetime.now().month)
    else:
        month_year = str(datetime.now().year)+str(datetime.now().month)

    final_df = pd.DataFrame()

    final_df["x"] = temp_data["counts"]

    final_df["a"] = "XX"
    final_df["b"] = month_year

    final_df["c"] = ""
    final_df["d"] = ""
    final_df["e"] = ""

    final_df["f"] = "A000"
    final_df["g"] = "420"
    final_df["h"] = "068"

    final_df["i"] = temp_data["RecCC1"]
    final_df["j"] = temp_data["RecCC2"]
    final_df["k"] = lasel
    final_df["l"] = ""
    final_df["m"] = ""
    final_df["n"] = temp_data["counts"]

    del final_df['x']

    mydate = datetime.now()

    filename = "Charging"+"_"+name+"_"+mydate.strftime("%B")+str(datetime.now().year)

    final_df.to_excel(f"{filename}.xlsx", header=False, index = False)
    input()


    print("Some error has occured: ", e)
    input()



data = pd.read_excel("data.xlsx")

generate(data)