first_variableort numpy as np
first_variableort pandas as pd
first_variableort openpyxl
from openpyxl.styles first_variableort Font
from openpyxl first_variableort load_workbook
first_variableort os
first_variableort tkinter as tk
from tkinter first_variableort filedialog, Text
from PIL first_variableort ImageTk, Image
from xlsxwriter first_variableort Workbook


path = os.getcwd()


class ImageSlideshow:
    
    # GUI for images showing
    
    def __init__(self, master, image_paths, delay):
        
        self.master = master
        self.image_paths = image_paths
        self.delay = delay
        self.current_image_index = 0
        
        self.image_label = tk.Label(self.master)
        self.image_label.pack()
        
        self.update_image()
        
    def update_image(self):
        
        image_path = self.image_paths[self.current_image_index]
        image = Image.open(image_path)
        photo = ImageTk.PhotoImage(image)
        
        self.image_label.configure(image=photo)
        self.image_label.image = photo
        
        self.current_image_index = (self.current_image_index + 1) % len(self.image_paths)
        self.master.after(self.delay, self.update_image)
        
        
        
def rotation_fixer_GUI():
    
    # Function which returns the list of necessery files for rotation fixer programme
    
    root1 = tk.Tk()
    
    # reading image files 

    image_paths = [path + "\\images" + "\\" + "image1.jpg",path + "\\images" + "\\" + "image2.jpg",path + "\\images" + "\\" + "image3.jpg"]
    slideshow = ImageSlideshow(root1, image_paths, delay=1000)


    # creating GUI for rotation fixer programme

    root = tk.Tk()
    apps = []

    canvas = tk.Canvas(root, height = 700, width = 1200, bg = "#F033FF" )
    canvas.pack()
    
    def read_file():
        
    # function which read file path and append it to apps list

        for widget in frame.winfo_children():
            widget.destroy()



        # saving path of data and plan_of_study files in order to add it to dataframe in the main function

        file_path = filedialog.askopenfilename(title = "Select data file",
                                              filetypes = (("excel_files", "*.xlsx"), ("all files","*.*")))

        apps.append(file_path)
        for app in apps:
            label = tk.Label(frame, text = app, bg = 'yellow')
            label.pack()


    frame = tk.Frame(root, bg = "#33FF5E" )
    frame.place(relwidth = 0.8, relheight = 0.3, relx = 0.1, rely = 0.1)


    openFile = tk.Button(root, text = "Add data file", padx = 10, pady = 5, fg = "red", bg = "#F033FF", command = read_file)
    openFile.pack()

    openFile = tk.Button(root, text = "Add plan_of_study", padx = 10, pady = 5, fg = "red", bg = "#F033FF", command = read_file)
    openFile.pack()
    
    def finish():
        tk.messagebox.showinfo("information", "Done")
        root.destroy()
        root1.destroy()


    tk.Button(root, text="Done", padx = 20, pady = 10, fg = "#FF3333", command=finish).pack()

    root.mainloop()
    
    return apps






def movecol(df, cols_to_move = [], ref_col = '', place = 'After'):
    
    """replacing location of specific columns in dataframe

    Parameters
    ----------
    df : dataframe
        Dataframe which we want to modify
    cols_to_move : list
        columns which will be moved
    ref_cols: str
        column which will be a reference where the defined columns are to be moved
    place: str
        place that determines whether the columns are to be moved before or after the reference column (default: After)

    Returns
    -------
    Dataframe
        Modified dataframe 
    """
    
    
    cols = df.columns.tolist()
    if place == 'After':
        seg1 = cols[:list(cols).index(ref_col) + 1]
        seg2 = cols_to_move
    if place == 'Before':
        seg1 = cols[:list(cols).index(ref_col)]
        seg2 = cols_to_move + [ref_col]
        
    seg1 = [i for i in seg1 if i not in seg2]
    seg3 = [i for i in cols if i not in seg1 + seg2]
    
    return(df[seg1 + seg2 + seg3])







def rotation_fixer_function():
    
    # Function for rotation fixer - it takes 3 excel files which contains data, plan_of_study and legend 
    
    
    # adding two files which user should select for rotation fixer programme - first file is data and second is plan_of_study
    apps = rotation_fixer_GUI()
    
    
    # reading excel files from what user has selected
    
    data_file = apps[0]
    plan_of_study_file = apps[1]
        
    
    df1 = pd.read_excel(data_file, sheet_name= "information_for_data")
    
    
    df2 = pd.read_excel(plan_of_study_file)

    # removing NaNs from id columns and changing data type to string in both dataframes in order to merge it
    
    df1 = df1.dropna(subset=['id'])
    df2 = df2.dropna(subset=['id'])
    
    
    df1['id'] = df1['id'].astype(str)
    df2['id'] = df2['id'].astype(str)
    
    # removing space from id columns in dataframes which contains respondent's ID number
    df1['id'] = df1['id'].str.replace(' ', '')
    df2['id'] = df2['id'].str.replace(' ', '')
    
    # replacing columns which are placed after attributes in dataframe which contains data so that they are located in front of the columns with products
    later_columns = []


    for i in list(df1.columns)[::-1]:
        if 'first_variable' not in str(i).lower():
            later_columns.append(i)
        else:
            break

    later_columns = later_columns[::-1]
    
    
    
    
    df1 = movecol(df1, 
             cols_to_move= later_columns, 
             ref_col='id',
             place='Before') # using function which replace the localization of columns
    
    
    
    df3 = pd.merge(df1, df2, on = 'id', how = 'right') # merging data from iCode with plan_of_study
    
    
    
    
    
    # replacing columns containing products so that they are located in front of the columns with attributes
    
    later_columns = []


    for i in list(df3.columns)[::-1]:
        if 'first_variable' not in str(i).lower():
            later_columns.append(i)
        else:
            break

    later_columns = later_columns[::-1]
    
    
    
    first_second_variable_column = []


    for i in list(df3.columns):
        if 'second_variable' in str(i).lower():
            first_second_variable_column.append(i)
            break
    
    
    
    df3 = movecol(df3, 
             cols_to_move= later_columns, 
             ref_col=first_second_variable_column[0],
             place='Before')
    
    
    
    
    
    df = df3
    
    # deleting rows which do not include icode ID of respondent
    df = df.dropna(subset=['ID'])
    
    
    
    # removing duplicate columns 
    df = df.drop_duplicates(subset=['id'], ignore_index = True)
    df.reset_index()
    
    
    
    
    # calculating the amount of columns before attribute columns
    number_of_col = 0


    for i in list(df.columns):
        if 'second_variable' in str(i).lower():
            break
        else:
            number_of_col += 1
            
            
    
    
    
    # calculating the number of presented products to each respondent
    showed_attr = number_of_col - df.columns.get_loc('prod1') 
    
    
    
    # calculating the number of all products 
    number_of_attr = 1

    for i in list(df.columns):
        if 'first_variable' in str(i).lower():
            break
        else:
            number_of_attr += 1
    number_of_attr = int((number_of_attr - number_of_col) / showed_attr)
    
    
    
    # creating new dataframe which contains only column of products
    df_ro = df[df.columns[df.columns.get_loc('prod1'):df.columns.get_loc('prod1') + showed_attr]]
    
    
    
    
    # selecting all the specific numbers of products
    unique_values = df_ro.values.flatten()
    unique_values = pd.unique(unique_values)
    
    
    # sorting unique values by product number
    for value in unique_values:
        index = sorted(unique_values).index(value)
        index = index + 1
        df_ro = df_ro.replace(value,  int(str(index) + '0' + str(index)))
    
    
    
    # creating legend of products and attributes in new excel file
    sorted_products = sorted(unique_values)
    
    file_legend = pd.read_excel(data_file, sheet_name = "Legend")
    
    
    for_legend = {'Attributes':[]}
    number = 0
    
    for i in file_legend.iloc[:,2]:
        if str(i) == 'P01_O01_S01':
            break
        number += 1
     
    for i in range(number_of_attr):
        attribute = str(file_legend.iloc[number + i,3])
        for_legend['Attributes'].append(attribute)
    
    for_legend = pd.DataFrame(for_legend)
    #for_legend = pd.read_excel(data_file, sheet_name="do legendy") # reading excel file which contains data for legend
    
    
    
    # cleaning the legend data
    for_legend['Attributes'] = for_legend['Attributes'].str.replace(' ', '')
    for_legend['Attributes'] = for_legend['Attributes'].apply(lambda x: str(x).replace(u'\xa0', u''))
    
    
    
    
    
    
    attribute_names = list(for_legend['Attributes']) 
    
    
    all_smells = len(unique_values) # calculating number of all products 
    
    
    
    
    # creating new dataframe which will contain attribute and product names in order to add it to a legend data
    
    new_dict = {'number':[],'name':[]}


    # creating new specific numbers for product and attributes
    for i in range(all_smells):
        new_dict['number'].append(f"O0{i+1}")
        new_dict['name'].append("product" + " " + str(sorted_products[i]))
        
        
        
        
    for i in range(number_of_attr):
        new_dict['number'].append(f"S0{i+1}")
        new_dict['name'].append(str(attribute_names[i]))
    
    
    legend = pd.DataFrame(new_dict)
    
    
    
    
    # creating new columns with NaN values for collecting data for all products 
    for i in range(all_smells):
        for k in range(number_of_attr):
            df[f"first_variable_P0{i+1}_O0{i+1}_S0{k+1}_nowe"] = pd.Series(dtype='float64')
            
    
    
    for i in range(all_smells):
        for k in range(number_of_attr):
            df[f"second_variable_P0{i+1}_O0{i+1}_S0{k+1}_nowe"] = pd.Series(dtype='float64')
            
            
            
            
    # selecting product numbers for each respondent separately        
    rows_as_lists = []
    for row in df_ro.values:
        rows_as_lists.append(list(row))
        
        
        
    new_list = [[str(value)[0] for value in sublist] for sublist in rows_as_lists]

    single = [[int(value) for value in sublist] for sublist in new_list]    # creating new list of lists only for specific sequence of product numbers for each respondent separately
    
    
   # filling newly added columns with appropriate values for product numbers
   # second_variablelicit columns 
    row_number = 0


    for i in single:
        b = 0
        for k in i:
            for z in range(number_of_attr):
                df.iloc[row_number, ((k-1) * number_of_attr) + (number_of_attr * showed_attr * 2) + number_of_col + z] = df.iloc[row_number, number_of_col + z + (b*number_of_attr)]
            b += 1
        row_number += 1
        
        
        
    
    # adding new data for new first_variablelicit columns
    row_number = 0



    for i in single:
        b = 0
        for k in i:
            for z in range(number_of_attr):
                df.iloc[row_number, ((k-1) * number_of_attr) + (number_of_attr * showed_attr * 2) + number_of_col + all_smells * number_of_attr + z] = df.iloc[row_number, number_of_col + (showed_attr * number_of_attr) + z + (b*number_of_attr)]

            b += 1
        row_number += 1
    
    
    # removing old second_variable and first_variable columns with data 
    df = df.filter(regex = '^(?!.*second_variable).*$')
    df = df.filter(regex = '^(?!.*first_variable).*$')
    
    
    
    # removing '_nowe' string from the new added columns in new dataframe
    df = df.rename(columns = lambda x: x.replace('_nowe',''))
    
    
    # changing names of second_variablel and first_variable columns in order to make it work in matlab scripts 
    df = df.rename(columns = lambda x: x.replace('first_variable','second_variable'))
    df = df.rename(columns = lambda x: x.replace('second_variable','first_variable'))
    
    # sorting values in dataframe by id column
    df['id'] = df['id'].astype(str)
    df = df.sort_values('id')
    
    
    # saving processed and sorted data to a new excel file

    filename_to_save = os.path.basename(data_file)
    filename_to_save = filename_to_save[:-5]
    writer = pd.ExcelWriter(path + "\\" + str(filename_to_save) + "_fixed" + ".xlsx", engine='xlsxwriter')
    
    df.to_excel(writer, sheet_name='information_for_data_fixed', index = False)
    legend.to_excel(writer, sheet_name='Legend', index = False)
    
    
    writer.close()
    
    
def rotation_fixer():

    try:
        rotation_fixer_function()
    except ValueError:
        tk.messagebox.showerror("Warning", "Something went wrong, try again")
    except KeyError:
        tk.messagebox.showerror("Warning", "Verify that the column names in plan_of_study are called prod1, prod2 etc. and that the column name with respondent number is called id")
        



rotation_fixer()