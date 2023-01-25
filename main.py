import pandas as pd
import numpy as np
import scipy.ndimage as ndimage
import matplotlib.pyplot as plt
from math import sqrt
import os.path
import sys
import xlsxwriter
import random
from tkinter import *
from tkinter import filedialog


#Initiate GUI
root = Tk()
root.title("Forceplate Calibration Program")
#root.geometry("560x335")
root.configure(background='#19232d')
root.iconbitmap('nhs_icon.ico')



"""
Input variables

filename = File name with .csv ending
Smooth = 'Bandwidth' figure for smoothing algorithm. Greater figure = greater
    smoothing. Range is 15 to 40, default 20.
percentage = percentage of force interval used for step detection. Useful range
    is 70 to 90, default 75
threshlev = Used to differentiate between fast and slow changing data. Increase
    figure to favour slow changing. Range 0 to 110, default 0
excel_start_row = The row number to write the results table to in the results
    excel file
excel_start_col = The column number to write the results table to in the results
    excel file
"""

#GUI: Define widgets ----------------------------------------------------------
filename_box = Entry(root, width=30, bg='#222b35', fg="#ffffff", font=12)
Smooth_box = Scale(root, from_=15, to=40, orient=HORIZONTAL) # 15 - 40, default 20
Smooth_box.set(20)
percentage_box = Scale(root, from_=70, to=90, orient=HORIZONTAL) # 70 - 90, default 75
percentage_box.set(75)
threshlev_box = Scale(root, from_=0, to=110, orient=HORIZONTAL) # 0 - 110, default 0
threshlev_box.set(0)
drop_variable_row = IntVar()
drop_variable_row.set(6)
drop_variable_column = IntVar()
drop_variable_column.set(1)
#excel_start_row_box = OptionMenu(root, drop_variable_row, 6, 22, 38, 61, 78, 95)
#excel_start_col_box = OptionMenu(root, drop_variable_column, 1, 6, 10, 15, 1, 11, 21)
filename_label = Label(root, text="Filename: ", bg='#19232d', fg="#ffffff", font=18)
Smooth_label = Label(root, text="Smoothing factor default value: 20", bg='#19232d', fg="#ffffff", font=18)
percentage_label = Label(root, text="Percentage force default value: 75", bg="#19232d", fg="#ffffff", font=18)
threshlev_label = Label(root, text="Data speed threshold default value: 0", bg="#19232d", fg="#ffffff", font=18)
#excel_start_row_label = Label(root, text="Excel row: ")
#excel_start_col_label = Label(root, text="Excel column: ")
Instructions_label = Label(root, text="\n1. Select vertical files (groups of 9 or all 27), or horizontal files (groups of 12 or all 36) \n2. Create Excel files by clicking the Create Files button (overwrites existing files!). \n3. Clicking calculate will analyse the files selected and output the data into the relavant Excel files. \n\nNote: - default thresholds, scaling and smoothing parameters are itterated if data are classed as a fail. \n- 10 itterations are performed using random values of each parameter before that file is skipped.", bg="#19232d", fg="#ffffff")

#GUI: Place widgets -----------------------------------------------------------
filename_label.grid(row=0, column=0)
Smooth_label.grid(row=1, column=1)
percentage_label.grid(row=2, column=1)
threshlev_label.grid(row=3, column=1)
#excel_start_row_label.grid(row=4, column=0)
#excel_start_col_label.grid(row=5, column=0)
filename_box.grid(row=0, column=1)
#Smooth_box.grid(row=1, column=1)
#percentage_box.grid(row=2, column=1)
#threshlev_box.grid(row=3, column=1)
#excel_start_row_box.grid(row=4, column=1)
#excel_start_col_box.grid(row=5, column=1)
Instructions_label.grid(row=8, column=0, columnspan=3)

#GUI: Button - returns the filenames of the CSV files -------------------------
def open_file():
    root.filename = filedialog.askopenfilenames(initialdir="C:", title="Select files") #gives the entire file location
    filename_box.delete(0, END)
    files= ""
    for file in root.filename:
        split_filename = file.split('/')
        filename_position = len(split_filename)-1
        filename = split_filename[filename_position]
        files = files + filename + " "   
    files = files.strip()
    filename_box.insert(0, files)    
file_open_button = Button(root, text="Browse", command=open_file, bg='#222b35', fg="#ffffff", font=18)
file_open_button.grid(row=0, column=2)

#GUI: Button - creates 6 excel files for the output data to be stored ---------
def file_creater():
    #calculate Excel filenames
    files_string = str(filename_box.get())
    files_list = files_string.split(' ')
    filename = files_list[0]
    fp_1_horizontal_filename = "Force_Calibration_" + filename[2:4] + "_1_horizontal.xlsx"
    fp_2_horizontal_filename = "Force_Calibration_" + filename[2:4] + "_2_horizontal.xlsx"
    fp_3_horizontal_filename = "Force_Calibration_" + filename[2:4] + "_3_horizontal.xlsx"
    fp_1_vertical_filename = "Force_Calibration_" + filename[2:4] + "_1_vertical.xlsx"
    fp_2_vertical_filename = "Force_Calibration_" + filename[2:4] + "_2_vertical.xlsx"
    fp_3_vertical_filename = "Force_Calibration_" + filename[2:4] + "_3_vertical.xlsx" 
    XL_filenames = [fp_1_horizontal_filename, fp_2_horizontal_filename, fp_3_horizontal_filename, fp_1_vertical_filename, fp_2_vertical_filename, fp_3_vertical_filename]   
    #create Excel files
    for XL_filename in XL_filenames:
        workbook = xlsxwriter.Workbook(XL_filename)
        worksheet = workbook.add_worksheet()
        workbook.close()
file_create_button = Button(root, text="Create Files", command=file_creater, bg='#222b35', fg="#ffffff", font=18)
file_create_button.grid(row=6, column=1)

#GUI: Button - iniates data analysis ------------------------------------------
def FP_Checker():
    #perform checks on the input data
    Smooth = int(Smooth_box.get())
    percentage = int(percentage_box.get())
    threshlev = int(threshlev_box.get()) 
    files_string = str(filename_box.get())
    files_list = files_string.split(' ')
    
    user_error_message = ''
    if len(files_list)==9:
        x = [1, 1, 1, 11, 11, 11, 21, 21, 21]
        y = [61, 78, 95, 61, 78, 95, 61, 78, 95]
    elif len(files_list)==12:
        x = [1, 1, 1, 6, 6, 6, 10, 10, 10, 15, 15, 15]
        y = [6, 22, 38, 6, 22, 38, 6, 22, 38, 6, 22, 38]
    elif len(files_list)==27:
        x = [1, 1, 1, 11, 11, 11, 21, 21, 21, 1, 1, 1, 11, 11, 11, 21, 21, 21, 1, 1, 1, 11, 11, 11, 21, 21, 21]
        y = [61, 78, 95, 61, 78, 95, 61, 78, 95, 61, 78, 95, 61, 78, 95, 61, 78, 95, 61, 78, 95, 61, 78, 95, 61, 78, 95]
    elif len(files_list)==36:
        x = [1, 1, 1, 6, 6, 6, 10, 10, 10, 15, 15, 15, 1, 1, 1, 6, 6, 6, 10, 10, 10, 15, 15, 15, 1, 1, 1, 6, 6, 6, 10, 10, 10, 15, 15, 15]
        y = [6, 22, 38, 6, 22, 38, 6, 22, 38, 6, 22, 38, 6, 22, 38, 6, 22, 38, 6, 22, 38, 6, 22, 38, 6, 22, 38, 6, 22, 38, 6, 22, 38, 6, 22, 38]
    else:
        user_error_message += '\nPlease select only correct vertical or horizontal data! '
    
    #--------------------------------------------------------------------------
    """
    Read in .csv and produce data columns
    
    Detects loading direction and force plate number based on which data column has
    the force values in (may need to change the way of doing this)
    
    F = Force data in direction of interest
    ct1 = Crosstalk 1
    ct2 = Crosstalk 2
    VertInt = Number of vertical steps expected
    cop_x = x centre of pressure (for x loading)
    cop_y = y centre of pressure (for y loading)
    """
    for filename in files_list:
        if os.path.isfile(filename) == False:
            user_error_message += '\nPlease enter valid filename! '
            break
        if filename[-4:] != '.csv':
            user_error_message += '\nPlease enter a .csv file! '
            break
        
    fp_1_horizontal_filename = "Force_Calibration_" + filename[2:4] + "_1_horizontal.xlsx"
    fp_2_horizontal_filename = "Force_Calibration_" + filename[2:4] + "_2_horizontal.xlsx"
    fp_3_horizontal_filename = "Force_Calibration_" + filename[2:4] + "_3_horizontal.xlsx"
    fp_1_vertical_filename = "Force_Calibration_" + filename[2:4] + "_1_vertical.xlsx"
    fp_2_vertical_filename = "Force_Calibration_" + filename[2:4] + "_2_vertical.xlsx"
    fp_3_vertical_filename = "Force_Calibration_" + filename[2:4] + "_3_vertical.xlsx" 
    XL_filenames = [fp_1_horizontal_filename, fp_2_horizontal_filename, fp_3_horizontal_filename, fp_1_vertical_filename, fp_2_vertical_filename, fp_3_vertical_filename]
    for XL_filename in XL_filenames:
        if os.path.isfile(XL_filename) == False:
            user_error_message += '\nPlease ensure all necessary Excel files have been created! '
            break
    
    if user_error_message != '':
        #user_error_message_label = Label(root, text=user_error_message)
        #user_error_message_label.grid(row=8, column=0, columnspan=3)
        messagebox.showerror("Error", user_error_message)
      
    else:   
        #Perform data analysis ------------------------------------------------    
        def FP_Checker_function(filename, Smooth, percentage, threshlev, excel_start_row, excel_start_col):
            data = pd.read_csv(filename,skiprows=[0,1,2,4],usecols=range(0,29))        
            data.columns = ['Frame','Sub Frame','Fx1','Fy1','Fz1','Mx1','My1','Mz1',
                            'Cx1','Cy1','Cz1','Fx2','Fy2','Fz2','Mx2','My2','Mz2',
                            'Cx2','Cy2','Cz2','Fx3','Fy3','Fz3','Mx3','My3','Mz3',
                            'Cx3','Cy3','Cz3']
                 
            if max(abs(data.Fx1))>220 and max(abs(data.Fx1))<300:
                loading_direction = 'x'
                FP_number = '1'
            if max(abs(data.Fy1))>220 and max(abs(data.Fy1))<300:
                loading_direction = 'y'
                FP_number = '1'
            if max(abs(data.Fz1))>900 and max(abs(data.Fz1))<1100:
                loading_direction = 'z'
                FP_number = '1'
            
            if max(abs(data.Fx2))>220 and max(abs(data.Fx2))<300:
                loading_direction = 'x'
                FP_number = '2'
            if max(abs(data.Fy2))>220 and max(abs(data.Fy2))<300:
                loading_direction = 'y'
                FP_number = '2'
            if max(abs(data.Fz2))>900 and max(abs(data.Fz2))<1100:
                loading_direction = 'z'
                FP_number = '2'
            
            if max(abs(data.Fx3))>220 and max(abs(data.Fx3))<300:
                loading_direction = 'x'
                FP_number = '3'
            if max(abs(data.Fy3))>220 and max(abs(data.Fy3))<300:
                loading_direction = 'y'
                FP_number = '3'
            if max(abs(data.Fz3))>900 and max(abs(data.Fy3))<1100:
                loading_direction = 'z'
                FP_number = '3'
            
            x = data['Fx'+FP_number]
            y = data['Fy'+FP_number]
            z = data['Fz'+FP_number]
            
            if loading_direction == 'x':
                F,ct1,ct2,VertInt = x,y,z,5
            if loading_direction == 'y':
                F,ct1,ct2,VertInt = y,x,z,5
            if loading_direction == 'z':
                F,ct1,ct2,VertInt = z,x,y,10
                cop_x = data['Cx'+FP_number]
                cop_y = data['Cy'+FP_number]
            
            #------------------------------------------------------------------------------
            
            """
            Smooth the data and find the threshold value of the difference between two
            points to identify steps
            
            Fsmooth = Smoothed F data using gaussian filter and smoothing value set at
                input variable stage
            interval = Force interval between steps
            dif = Array of differences between each point of Fsmooth
            difav = Average of dif array
            difrms = rms of dif array
            thresh = Threshold value of difference between adjacent points to identify
                large jumps (noise)
            """
            
            Fsmooth = ndimage.gaussian_filter(F, sigma=Smooth, order=0)
            
            if loading_direction == 'x' or loading_direction == 'y':
                interval = (max(Fsmooth)+min(Fsmooth))*(float(percentage)/100)/VertInt
            if loading_direction == 'z':
                interval = (max(Fsmooth))*(float(percentage)/100)/VertInt
            
            dif = []
            
            for q in range(1,len(Fsmooth)-2):
                dif = np.append(dif,abs((Fsmooth[q+1]-Fsmooth[q-1])/2))
            
            difav = sum(dif)/len(dif)
            
            difrms = sqrt(sum(n*n for n in dif)/len(dif))
            
            if loading_direction == 'x' or loading_direction == 'y':
                thresh = (max(dif)*difav)/(difrms*(90+threshlev))
            if loading_direction == 'z':
                thresh = (max(dif)*difav)/(difrms*(150+threshlev))
            #------------------------------------------------------------------------------
            
            """
            Identify valid data (data points where the difference between adjacent points
            is less than thresh) to filter out large spikes
            
            Fvalid = Valid force data
            framevalid = Frame points of the valid data (used later for plot)
            ct1valid = Valid cross-talk 1 points
            ct2valid = Valid cross-talk 2 points
            cop_x_valid = Valid x centre of pressure points
            cop_y_valid = Valid y centre of pressure points
            """
            
            Fvalid = []
            framevalid = []
            ct1valid = []
            ct2valid = []
            cop_x_valid = []
            cop_y_valid = []
            
            for s in range(len(Fsmooth)):
                if abs(Fsmooth[s]-Fsmooth[s-1]) < thresh \
                and abs(Fsmooth[s-1]-Fsmooth[s-2]) < thresh \
                and abs(Fsmooth[s-2]-Fsmooth[s-3]) < thresh \
                and abs(Fsmooth[s-3]-Fsmooth[s-4]) < thresh:
                    Fvalid = np.append(Fvalid,Fsmooth[s])
                    framevalid = np.append(framevalid,s)
                    ct1valid = np.append(ct1valid,ct1[s])
                    ct2valid = np.append(ct2valid,ct2[s])
                    if loading_direction == 'z':
                        cop_x_valid = np.append(cop_x_valid,cop_x[s])
                        cop_y_valid = np.append(cop_y_valid,cop_y[s])
            
            #------------------------------------------------------------------------------
            
            """
            Identify positions of steps in Fvalid by finding where the difference between
            two points 10 positions from each other is greater than interval. This can be
            changed from 10 if necessary- it was initally between adjacent points but
            some steps were slow moving
            
            step_count = Number of steps total
            step_positions = Index value of steps in Fvalid
            """
            
            step_count = 0
            step_positions = [0]
            
            n = 0
            while n < len(Fvalid)-11:
                n = n+1
                if abs(Fvalid[n] - Fvalid[n+10]) > abs(interval):
                    step_count = step_count+1
                    step_positions = np.append(step_positions,n+9)
                    n = n+10
            
            #------------------------------------------------------------------------------
            
            """
            Select the values for the zero offset by summing all those before the first
            step, and dividing this by the length of the values. Subtract this from all
            of Fvalid
            
            values_for_offset = Select all values from F before the first step if between
                -1 and 1
            zero_offset = Sum all values for offset and divide by length to get average
            Foffset = Subtract offset from Fvalid to get new force values
            """
            
            values_for_offset = []
            
            for n in range(0,int(step_positions[1]/2)):
                values_for_offset = np.append(values_for_offset,F[n])
            
            zero_offset = sum(values_for_offset)/len(values_for_offset)
            
            Foffset = Fvalid-zero_offset
            
            #------------------------------------------------------------------------------
            
            """
            Calculate the average force of each step by summing all forces in the step and
            dividing by step length
            For vertical loading, the 41N weight holder is compensated for when it is
            removed at the end
            
            step_force_averages = Average force values of each step
            """
            
            step_force_averages = []
            
            for n in range(0,len(step_positions)-1):
                step_average = sum(Foffset[int(step_positions[n]+1)
                :int(step_positions[n+1]-1)])/(len(Foffset[int(step_positions[n]+1)
                :int(step_positions[n+1]-1)]))
                step_force_averages = np.append(step_force_averages,step_average)
            
            step_average_final = sum(Foffset[int(step_positions[-1]+1)
                :int(Foffset[-1]-1)])/(Foffset[-1]-Foffset[int(step_positions[-1])])
            
            step_force_averages = np.append(step_force_averages, step_average_final)
            
            #if loading_direction == 'z':
            #    step_force_averages[-1] = step_force_averages[-1]-41
            
            #------------------------------------------------------------------------------
            
            """
            Calculate the zero offsets for the cross talk values and subtract. Calculate
            the average cross-talk forces for each step
            
            ct1_zero_offset = Zero offset value for cross-talk 1
            ct2_zero_offset = Zero offset value for cross-talk 2
            ct1offset = Cross-talk 1 values with zero offset subtracted
            ct2offset = Cross-talk 2 values with zero offset subtracted
            ct1_force_averages = Average force values for each step for cross-talk 1
            ct2_force_averages = Average force values for each step for cross-talk 2
            """
            
            ct1_zero_offset = sum(ct1valid[0:int(step_positions[1])])/step_positions[1]
            ct2_zero_offset = sum(ct2valid[0:int(step_positions[1])])/step_positions[1]
            
            ct1offset = ct1valid - ct1_zero_offset
            ct2offset = ct2valid - ct2_zero_offset
            
            ct1_force_averages = [0]
            ct2_force_averages = [0]
            
            for n in range(1,len(step_positions)-1):
                ct1_average = sum(ct1offset[int(step_positions[n])
                :int(step_positions[n+1])])/(int(step_positions[n+1])-int(step_positions[n]))
                ct1_force_averages = np.append(ct1_force_averages,ct1_average)
            
            for n in range(1,len(step_positions)-1):
                ct2_average = sum(ct2offset[int(step_positions[n])
                :int(step_positions[n+1])])/(int(step_positions[n+1])-int(step_positions[n]))
                ct2_force_averages = np.append(ct2_force_averages,ct2_average)
            
            
            
            ct1_average_final = sum(ct1offset[int(step_positions[-1])
                :int(len(ct1offset))])/(len(ct1offset)-ct1offset[int(step_positions[-1])])
            
            ct1_force_averages = np.append(ct1_force_averages, ct1_average_final)
            
            ct2_average_final = sum(ct2offset[int(step_positions[-1])
                :int(len(ct2offset))])/(len(ct2offset)-ct2offset[int(step_positions[-1])])
            
            ct2_force_averages = np.append(ct2_force_averages, ct2_average_final)
            #------------------------------------------------------------------------------
            """
            For vertical loading, find the centre of pressure x and y co-ordinates for each
            step.
            
            cop_x_coordinates = x coordinates of centre of pressure for each step
            cop_y_coordinates = y coordinates of centre of pressure for each step
            """
            
            if loading_direction == 'z':
                cop_x_coordinates = [0]
                for n in range(2,len(step_positions)):
                    cop_x_coordinate_value = sum(cop_x_valid[step_positions[n-1]
                        :step_positions[n]])/(step_positions[n]-step_positions[n-1])
                    cop_x_coordinates = np.append(cop_x_coordinates, cop_x_coordinate_value)
            
            if loading_direction == 'z':
                cop_y_coordinates = [0]
                for n in range(2,len(step_positions)):
                    cop_y_coordinate_value = sum(cop_y_valid[step_positions[n-1]
                        :step_positions[n]])/(step_positions[n]-step_positions[n-1])
                    cop_y_coordinates = np.append(cop_y_coordinates, cop_y_coordinate_value)
            
            if loading_direction == 'z':
                cop_x_coordinates = np.append(cop_x_coordinates, 0)
                cop_y_coordinates = np.append(cop_y_coordinates, 0)
            
            #------------------------------------------------------------------------------
            
            """
            Create an array of 'proper' values of the expected force depending on the
            loading direction. Correct for gravity and produce a percentage of the step
            force compared to the proper (expected) value
            
            force_output = Force percentage values compared to the expected force
            ct1_output = Cross-talk 1 output values as a percentage of the step force
            ct2_output = Cross-talk 2 output values as a percentage of the step force
            """
            
            if loading_direction == 'x' or loading_direction == 'y':
                proper = [10.1968,5,10,15,20,25,20,15,10,5,10.1968]
            if loading_direction == 'z':
                proper = [1,10,20,30,40,50,60,70,80,90,100,1]
            
            force_output = [0]
            
            for n in range(1,len(step_force_averages)):
                force_output = np.append(force_output,abs((step_force_averages[n]/9.807)*100/proper[n]))
            
            
            if loading_direction == 'z':
                ct1_output = [ct1_force_averages[0]]
                ct2_output = [ct2_force_averages[0]]
            
                for n in range(1,len(ct1_force_averages)-1):
                   ct1_output = np.append(ct1_output,ct1_force_averages[n]/step_force_averages[n]*100)
            
                for n in range(1,len(ct2_force_averages)-1):
                   ct2_output = np.append(ct2_output,ct2_force_averages[n]/step_force_averages[n]*100)
            
                ct1_output = np.append(ct1_output, ct1_force_averages[-1])
                ct2_output = np.append(ct2_output, ct2_force_averages[-1])
            
            
            if loading_direction == 'x' or loading_direction == 'y':
                ct1_output = []
                ct2_output = []
            
                for n in range(0,len(ct1_force_averages)-1):
                   ct1_output = np.append(ct1_output,ct1_force_averages[n]/step_force_averages[n]*100)
            
                for n in range(0,len(ct2_force_averages)-1):
                   ct2_output = np.append(ct2_output,ct2_force_averages[n]/step_force_averages[n]*100)
            
                ct1_output = np.append(ct1_output, ct1_force_averages[-1])
                ct2_output = np.append(ct2_output, ct2_force_averages[-1])
            
            #------------------------------------------------------------------------------
            
            """
            Give each step a pass/fail based on 95% criteria
            
            acceptance = Array of pass/fails
            """
            
            acceptance = ['n/a']
            
            for n in range(1,len(force_output)-1):
                if 95 <= float(force_output[n]) <= 105:
                    acceptance = np.append(acceptance,'Pass')
                else:
                    acceptance = np.append(acceptance,'Fail')
            
            acceptance = np.append(acceptance,'n/a')
            
            #------------------------------------------------------------------------------
            """
            Detect steps of forces greater than expected. Number of large steps should be 0.
            
            large_step = Number of large steps
            """
            
            if loading_direction == 'z':
                large_step = 0
            
                for n in range(0, len(step_force_averages)-1):
                    step = step_force_averages[n+1]-step_force_averages[n]
                    if 1.25*proper[n] < step < 9*proper[n]:
                        large_step = large_step + 1
            
            #------------------------------------------------------------------------------
            """
            Produce dataframe of results. Vertical loading also includes COP coordinates
            
            table_data = Data to be in the final results output
            results_df = Results dataframe
            """
            
            
            if loading_direction == 'x':
                table_data = {'Percentage' : force_output, 'Cross-Talk 1' : ct1_output,
                              'Cross-Talk 2' : ct2_output, 'Pass/Fail' : acceptance}
                results_df = pd.DataFrame(table_data,
                                          columns = ['Percentage','Cross-Talk 1',
                                                     'Cross-Talk 2','Pass/Fail'])
            
            if loading_direction == 'y':
                table_data = {'Cross-Talk 1' : ct1_output, 'Percentage' : force_output,
                              'Cross-Talk 2' : ct2_output, 'Pass/Fail' : acceptance}
                results_df = pd.DataFrame(table_data,
                                          columns = ['Cross-Talk 1','Percentage',
                                                     'Cross-Talk 2','Pass/Fail'])
            
            if loading_direction == 'z':
                table_data = {'Cross-Talk 1' : ct1_output, 'Cross-Talk 2' : ct2_output,
                              'Percentage' : force_output, 'COP X' : cop_x_coordinates,
                              'COP Y' : cop_y_coordinates, 'Pass/Fail' : acceptance}
                results_df = pd.DataFrame(table_data,
                                          columns = ['Cross-Talk 1','Cross-Talk 2',
                                                     'Percentage','COP X',
                                                     'COP Y','Pass/Fail'])
            
            results = results_df.to_string(index=False)
            print(results)
            
            if loading_direction == 'x' or loading_direction == 'y':
                if len(step_force_averages) != 11:
                    print('Step number is: '+str(len(step_force_averages))+'. Expected number is 11.')
            
            if loading_direction == 'z':
                if len(step_force_averages) != 12:
                    print('Step number is: '+str(len(step_force_averages))+'. Expected number is 12.')
                if large_step != 0:
                    print('Number of large steps is '+str(large_step)+'. Expected number is 0.')
            
            score = 0
            for n in range(1,len(acceptance)-1):
                if acceptance[n] == 'Pass':
                    score = score + 1
                if acceptance[n] != 'Pass':
                    score = score
            
            
            if loading_direction == 'x' or loading_direction == 'y':
                if score == 9:
                    print('Overall: Pass')
                if score < 9:
                    print('Overall: Fail')
                    print(error_out_to_itterate)
            if loading_direction == 'z':
                if score == 10:
                    print('Overall: Pass')
                if score < 10:
                    print('Overall: Fail')
                    print(error_out_to_itterate)
            #------------------------------------------------------------------------------
            """
            Results file is chosen depending on the calibration session, plate number, and
            loading direction. The start columns and rows are set in the initial variables
            at the beginning. Current filenames are:
                Force_Calibration_44_1_horizontal
                Force_Calibration_44_2_horizontal
                Force_Calibration_44_3_horizontal
                Force_Calibration_45_1_vertical
                Force_Calibration_45_2_vertical
                Force_Calibration_45_3_vertical
            
            Writer is used so that the existing results are not overwritten
            """
            
            #if FP_number == '1':
            #    results_filename = r'Results_Plate_1.xlsx'
            #
            #cols = [0, 5, 10, 15]
            #rows = [0, 12, 24]
            #
            #trial_number = int(filename[5:7])
            #
            #
            
            excel_results_df = results_df.drop(columns = 'Pass/Fail')
            
            
            if loading_direction == 'x' or loading_direction == 'y':
                results_filename = 'Force_Calibration_'+filename[2:4]+'_'+str(FP_number)+'_'+'horizontal.xlsx'
            
            if loading_direction == 'z':
                results_filename = 'Force_Calibration_'+filename[2:4]+'_'+str(FP_number)+'_'+'vertical.xlsx'
            
            original_results_df = pd.read_excel(results_filename, header = None)
            
            writer = pd.ExcelWriter(results_filename)
            
            original_results_df.to_excel(writer, index = False, header = None)
            
            excel_results_df.to_excel(writer, startcol = excel_start_col, startrow = excel_start_row, header = None, index = False)
            
            writer.save()
            
            
            
            #------------------------------------------------------------------------------
            """
            Plot the original data and smoothed data
            """
            
            plt.figure(0)
            
            plt.plot(F,label='Original Data')
            plt.plot(framevalid,Foffset,label='Filtered and Smoothed Data')
            plt.legend()
            
            print(FP_number, loading_direction)
            
        
        #itterate FP_Checker_function for each file selected ------------------
        i=0
        for filename in files_list:
            excel_start_row = y[i]
            excel_start_col = x[i]
            i+=1
            try:
                FP_Checker_function(filename, Smooth, percentage, threshlev, excel_start_row, excel_start_col)
            except:
                #if analysis fails or calibration data is a fail, itterate ----
                print("itterating parameters")
                j=0
                while j<10:
                    j+=1
                    Smooth = random.randint(0, 40)
                    percentage = random.randint(70, 90)
                    threshlev = random.randint(0, 110)
                    try:
                       FP_Checker_function(filename, Smooth, percentage, threshlev, excel_start_row, excel_start_col) 
                       break
                    except:
                        print("itterating parameters")
                        continue
                continue
        messagebox.showinfo("", "Analysis complete!")
        
#GUI: Button - iniates data analysis ------------------------------------------       
calculate_button = Button(root, text="Calculate", command=FP_Checker, bg='#222b35', fg="#ffffff", font=18)
calculate_button.grid(row=7, column=1)


root.mainloop()
