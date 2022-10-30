import sys
import pandas as pd
import numpy as np
import random

def sort_v2(file_path,group_size):
    file = pd.read_excel(file_path)
    group_size = int(group_size)
    df2 = file.drop(file[file["# Responses"] == 0.0].index)

    #sort as ["Name","Study line","Type"]
    New_list = []
    new_row = None
    for row in range(df2.shape[0]):
        if pd.isnull(df2.iloc[row,0]):
            new_row.append(df2.iloc[row,7])
        else:
            if new_row != None:
                New_list.append(new_row)
            new_row = [df2.iloc[row,0]]
            
    results_df = pd.DataFrame(New_list)
    results_df.columns = ["Name","Study line","Type"]

    #adding studyline no and type no
    study_line_list = list(results_df["Study line"].unique())
    type_list = list(results_df["Type"].unique())
    results_df['Studyline_No']=results_df.apply(lambda x: study_line_list.index(x[1])+1,axis=1)
    results_df['type_No']=results_df.apply(lambda x: type_list.index(x[2])+1,axis=1)

    results_df_new = results_df.reset_index()
    gb = results_df_new.groupby('type_No').agg(lambda x: list(x))[['Studyline_No','index']]

    # Grouping 
    gb_copy = gb.copy()
    all_group = [] #store index
    group = [] #store index
    group_help = [] # just helping during the sort. store study line 
    all_group_type_sl = []  #all group type study line, store study line.
    group_type_sl = [] # singe group type study line

    limits = len(study_line_list) // group_size - 2 # limits for the group of same study line.  

    it_label = 0
    loe = range(len(type_list)) #left over elements

    while loe != 0:
        for x in range(len(type_list)):
            c = 0
            add = True

            if len(gb_copy.iloc[x][0]) ==0: continue
            
            if len(group) >= group_size:
                all_group.append(group) 
                all_group_type_sl.append(group_type_sl)
                group=[]
                group_help = []
                group_type_sl = []
            else:
                while add == True:
                    #if gb_copy.iloc[x][0][c] not in group_help or group_help.count(gb_copy.iloc[x][0][c])==1: #and gb_copy.iloc[x][1][c] not in group:
                    if group_help.count(gb_copy.iloc[x][0][c]) <= limits:
                        group.append(gb_copy.iloc[x][1][c])
                        group_help.append(gb_copy.iloc[x][0][c])
                        group_type_sl.append((x+1,gb_copy.iloc[x][0][c]))
                        gb_copy.iloc[x][0].pop(c)
                        gb_copy.iloc[x][1].pop(c)
                        add = False
                        it_label = 0
                    
                    else:
                        if c < len(gb_copy.iloc[x][0])-1:
                            c += 1
                        else:
                            add = False

        loe = 0 #left over elements
        for x in range(len(type_list)):
            loe += len(gb_copy.iloc[x][0])
        
        it_label += 1
        if it_label == 3: 
            limits += 1
            it_label = 0


    #add new group when group > 0.8 * group size
    if len(group) >= 0.8 * group_size:
        while len(group) < group_size:
            group.append(-1)
            group_type_sl.append(-1)
        all_group.append(group) 
        all_group_type_sl.append(group_type_sl)

    all_group_arr = np.array(all_group)

    def get_group(idx):
        group = np.where(all_group_arr ==idx)[0]
        if len(group) == 0: 
            return -1
        else:
            return int(group)+1
    #assign gourp

    results_df['Group'] = results_df.reset_index().apply(lambda x: get_group(x['index']),axis=1)

    #assign random group to left over
    n_left = results_df.query('Group==-1').shape[0]
    randlist = random.sample(range(1,all_group_arr.shape[0]),n_left)

    def assign_rand_group(arrLike):
        if arrLike['Group'] == -1:
            k = randlist[-1]
            randlist.pop()
        else:
            k = arrLike['Group']
        return k

    results_df['desc'] = results_df.apply(lambda x: 'left' if x['Group'] == -1 else '', axis=1) #set label for random students
    #results_df['Group'] = results_df.apply(lambda x: randlist[0] if x['Group'] == -1 else x['Group'], axis=1)
    results_df['Group'] = results_df.apply(assign_rand_group, axis=1)


    #Summary Sheet
    Summary = pd.DataFrame({'Group':[i for i in range(1,results_df['Group'].max()+1)],'Group size':'=COUNTIF(Group!F:F,@Summary!A:A)'})

    writer = pd.ExcelWriter('results_sl.xlsx', engine='openpyxl')
    
    Summary.to_excel(writer, index=False,sheet_name = 'Summary')
    results_df.sort_values('Group').to_excel(writer, index=False, sheet_name='Group')

    writer.save()
    writer.close()
    return 

if __name__ == "__main__":
    file_path = sys.argv[1]
    group_size = int(sys.argv[2])
    
    sort_v2(file_path,group_size)
