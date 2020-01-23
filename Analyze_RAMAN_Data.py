"""
Name:           Rennishaw RAMAN Analyzer
Description:    This code anaylzes bulk SPC data files of a RAMAN run, exporting to CSV and providing images of averaged RAMAN with SEM considered
Author:         John Ferrier
Notes:          1. This code is a modified version of "convert_gui.py" by Rohan Isaac
                2. Base file names must follow a specific format such that each measured grain is indicated, along with how many measurements are made on that grain.
                    For Example:
                        SomeFileName_GrainNumber_MeasurementNumberOnGrain.spc

                        5 measurements on a single grain would look like this:
                            SomeFileName_1_1.spc
                            SomeFileName_1_2.spc
                            SomeFileName_1_3.spc
                            SomeFileName_1_4.spc
                            SomeFileName_1_5.spc

                        2 measurements each made on 3 grains would look like this:
                            SomeFileName_1_1.spc
                            SomeFileName_1_2.spc
                            SomeFileName_2_1.spc
                            SomeFileName_2_2.spc
                            SomeFileName_3_1.spc
                            SomeFileName_3_2.spc

                    This format is important for the code to differentiate which items should be compared for Averaging and SEM purposes
"""

#Import Libraries
from __future__ import division, absolute_import, unicode_literals, print_function
from tkinter import *
#from tkinter import Tk, StringVar, DISABLED, NORMAL, END, W, E, N, S, Menu
from tkinter.ttk import Frame, Label, Button, Radiobutton, Entry, Progressbar
from tkinter import filedialog
import spc
import os
import math
import numpy as np
import matplotlib.patches as mpatches
import matplotlib.pyplot as plt

class AnalyzeSPC:
    def __init__( self, master ):
        
        #Set Master TK Window from intiation (root)
        self.master             = master

        #Name the Window
        master.title("Analyze Rennishaw SPC files")

        #Style a general outside padding of 10px
        self.mf                      = Frame( master, padding = "10" )

        #Initiate Menubar
        self.menubar            = Menu( self.master )

        #File Menu
        self.filemenu           = Menu( self.menubar, tearoff = 0 )
        self.filemenu.add_command( label = "Exit", command = self.master.quit, accelerator = "Alt+F4"  )
        self.menubar.add_cascade(  label = "File", menu = self.filemenu )

        #Help Menu
        self.helpmenu           = Menu( self.menubar, tearoff = 0 )
        self.helpmenu.add_command( label = "How to use", command = lambda: self.output_message("How to use this application", "Base file names must follow a specific format such that each measured grain is indicated, along with how many measurements are made on that grain.\n\n\tFor Example:\n\n\t\tSomeFileName_GrainNumber_MeasurementNumberOnGrain.spc\n\n\t\t5 measurements on a single grain would look like this:\n\t\t\tSomeFileName_1_1.spc\n\t\t\tSomeFileName_1_2.spc\n\t\t\tSomeFileName_1_3.spc\n\t\t\tSomeFileName_1_4.spc\n\t\t\tSomeFileName_1_5.spc\n\n\t\t2 measurements each made on 3 grains would look like this:\n\t\t\tSomeFileName_1_1.spc\n\t\t\tSomeFileName_1_2.spc\n\t\t\tSomeFileName_2_1.spc\n\t\t\tSomeFileName_2_2.spc\n\t\t\tSomeFileName_3_1.spc\n\t\t\tSomeFileName_3_2.spc\n\n\tThis format is important for the code to differentiate which items should be compared for Averaging and SEM purposes.", "Close", None))

        self.menubar.add_cascade(  label = "Help", menu = self.helpmenu )

        self.master.config( menu = self.helpmenu )

        #Initiate Grid format for TKinter
        self.mf.grid(column = 0, row = 0, sticky = ( N, W, E, S ) )
        self.mf.columnconfigure( 0, weight = 1 )
        self.mf.rowconfigure(    0, weight = 1 )

        #Build Input variables
        self.message            = "Select folder containing *.SPC files to be Analyzed"
        self.label_text         = StringVar()
        self.folder             = StringVar()
        self.output_fmt         = StringVar( value = "csv")
        self.output_folder      = StringVar()
        self.output_plots       = IntVar( value = 1 )
        self.output_fmt_bool    = IntVar( value = 0 )
        self.progress_var       = DoubleVar( )

        #Set Initial Output Information for User
        self.label_text.set( self.message )

        #Label User Choices
        self.label              = Label( self.mf, textvariable = self.label_text )
        self.folder_label       = Label( self.mf, text = "Input Folder" )
        self.output_fmt_label   = Label( self.mf, text = "Output Format" )
        self.output_fold_label  = Label( self.mf, text = "Output Folder" )

        #Build Output Selection Choices for User
        self.fmt_txt            = Radiobutton( self.mf, text = "TXT", variable = self.output_fmt, value = 'txt', state = DISABLED )
        self.fmt_csv            = Radiobutton( self.mf, text = "CSV", variable = self.output_fmt, value = 'csv', state = DISABLED )
        self.fmt_plot           = Checkbutton( self.mf, text = "Plot Data", variable = self.output_plots, onvalue = 1, offvalue = 0 )
        self.fmt_bool           = Checkbutton( self.mf, text =  "Output Raw Data", variable = self.output_fmt_bool, onvalue = 1, offvalue = 0, command = self.output_enable)
        
        #Progressbar
        self.p_bar              = Progressbar( self.mf, length = 300, mode = 'determinate', variable = self.progress_var, maximum = 100. )

        #Build Directory Input/Output for User
        self.folder_entry       = Entry( self.mf, textvariable = self.folder )
        self.folder_output      = Entry( self.mf, textvariable = self.output_folder )

        #Build Buttons for User
        self.sel_folder         = Button( self.mf, text = "Browse",  command = self.ask_dir )
        self.sel_out_folder     = Button( self.mf, text = "Browse",  command = self.ask_out_dir )
        self.analyze_btn        = Button( self.mf, text = "Analyze", command = self.convert )

        #Build GUI Layout
        self.label.grid(             row = 0, column = 0, columnspan = 4, sticky = W + E)   #Label

        self.folder_label.grid(      row = 1, column = 0, sticky = E )                      #Input Folder Directory Label
        self.folder_entry.grid(      row = 1, column = 1, sticky = W + E, columnspan = 2 )  #Input Folder Input
        self.sel_folder.grid(        row = 1, column = 3, sticky = W )                      #Input Folder Directory Button

        self.fmt_bool.grid(          row = 2, column = 0, sticky = W, columnspan = 2 )      #Boolean that allows raw data output

        self.output_fmt_label.grid(  row = 3, column = 0, sticky = E )                      #Output Format type (TXT, CSV Radio buttons Label)
        self.fmt_txt.grid(           row = 3, column = 1, sticky = W )                      #Output Format type TXT button
        self.fmt_csv.grid(           row = 3, column = 2, sticky = W )                      #Output Format type CSV button
        self.fmt_plot.grid(          row = 3, column = 3, sticky = W )                      #Output Plots

        self.output_fold_label.grid( row = 4, column = 0, sticky = E )                      #Output Folder Directory Label
        self.folder_output.grid(     row = 4, column = 1, sticky = W + E, columnspan = 2 )  #Output Folder Input
        self.sel_out_folder.grid(    row = 4, column = 3, sticky = W )                      #Output Folder Directory Button

        self.analyze_btn.grid(       row = 5, column = 1, columnspan = 2, sticky = W + E )  #Analyze Files Button

        
        self.p_bar.grid(             row = 6, column = 0, columnspan = 4, sticky = E )

        #Add 5px padding to all elements
        for child in self.mf.winfo_children():
            child.grid_configure( padx = 5, pady = 5 )

    def output_enable(self):

        if not self.output_fmt_bool.get():
            self.fmt_txt.configure(state = DISABLED)
            self.fmt_csv.configure(state = DISABLED)
        else:
            
            self.fmt_txt.configure(state = NORMAL)
            self.fmt_csv.configure(state = NORMAL)

    def convert(self):

        #Get Folder names
        self.fol_val        = str( self.folder.get() )
        self.fol_out_val    = str( self.output_folder.get() )

        ffn_out             = os.path.abspath(self.fol_out_val)

        #Get output format (txt or csv)
        self.fmt_val        = str( self.output_fmt.get() )

        #Get whether to plot or not
        self.fmt_plot_val   = bool( self.output_plots.get() )

        #Inform user of starting conversions
        print( "Analyzing {} with {} ext".format(self.fol_val, self.fmt_val) )

        #Figure output file extension and delimiter
        if self.fmt_val == 'txt':
            exten = '.txt'
            delim = '\t'
        else:
            #Default is CSV
            exten = '.csv'
            delim = ','

        #Initiate grain and measurement indices
        self.grain          = 0
        self.measurement    = 0

        #Build list for files
        flist = []

        #Get file directory from folder value
        ffn = os.path.abspath( self.fol_val )

        #Check for output folder location. If none, default to Input folder
        out_str = self.fol_out_val.replace(" ","")

        #If No output folder selected
        if out_str == "":
            #Default to input folder
            self.fol_out_val = self.fol_val
            self.output_folder.set( value = self.fol_out_val )

        #Find how many grain files exist in folder
        self.grain_count = 0
        for f in os.listdir( ffn ):

            #Check extension
            if f.lower().endswith('spc'):

                #Split file name by underscores
                f_split = f.split("_")
                
                #Ensure correct format for tracking
                if len(f_split) >= 2:

                    #Look at second to last in split ( SomeFileName_1_1.spc )
                    if int( f_split[-2] ) > self.grain_count:
                        self.grain_count = int( f_split[-2] )

        #Consider grain count
        if self.grain_count == 0:
            self.output_message("Input Folder Empty", "Your input folder is either empty or the spc files are not configured correctly.", "Close", None)

        #Cycle through each grain count and append files
        for i in range(self.grain_count):

            #Build Measurements list
            mlist = []
            
            #Append sorted files to the filelist
            for f in os.listdir( ffn ):

                #Check extension
                if f.lower().endswith('spc'):
                    
                    fl = f

                    #Split file name to find current
                    f_split = fl.split("_")

                    #Ensure correct format for tracking
                    if len(f_split) >= 2:

                        #Check if current file is the current grain
                        if int( f_split[-2] ) == i+1:
                            #print( f )
                            #Add to measurement list
                            mlist.append( os.path.join( ffn , f) )

            #Append grain i+1 files
            flist.append( mlist )

        ######  STOPPED HERE! FILES SORTED BY GRAIN. START ANALYZING ######

        #Check if any output is selected. If not, make a silly message
        output_data = bool( self.output_fmt_bool.get() )

        output_plot = bool( self.output_plots.get() )

        if (not output_data) and (not output_plot):
            self.output_message( "What are you doing?", "... You have to select some output method...\n\nI mean, I could analyze this for you but you won't see it.", "Oh yeah, duh!", None  )

        #Else, continue analyzing
        else:
            prog_incr = 100./len( flist )

            #Look at grain i

            #Loop through i grains
            for i in range( len( flist ) ):                     #3

                tot_grain_data_list     = []                    #All unanalyzed grain data
                ave_grain_data_list     = []                    #Averaged data per grain
                sem_grain_data_list     = []                    #SEM data per grain

                #Loop through j measurements   
                for j in range( len( flist[i] ) ):              #5

                    #Open file into array
                    f       = spc.File( flist[i][j] )
                    xy_data = f.data_list()                     #xy data for grain i, measurement j

                    #Append collected data to the grain_data_list for later processing
                    tot_grain_data_list.append( xy_data )

                #print( len(tot_grain_data_list) )    #total measurements
                #print( len(tot_grain_data_list[0]) ) #x and y

                #Average arrays (function? OUT = Single array)
                normalized_to = 1.0

                ave_grain_data_list.append( self.grain_average( tot_grain_data_list, normalized_to ) )

                #Find SEM in arrays (function? OUT = Single array of +/-)
                sem_grain_data_list.append( self.SEM_calc( tot_grain_data_list, ave_grain_data_list[0] ) )

                #print( sem_grain_data_list[0][0][:5] )

                #if i==0:
                #    print( sem_grain_data_list[0][0][:10] )

                #If plot, save plot data as PNG
                if output_plot:
                    
                    #Plot normalized and averaged values
                    figsize = ( 16, 9 )
                    plt.figure( figsize = figsize, dpi = 600 )

                    plt.plot( ave_grain_data_list[0][0], ave_grain_data_list[0][1], color = '#000099', lw = 0.5 )

                    #Calc total SEM Values
                    sem_above = []
                    sem_below = []

                    

                    for k in range( len( sem_grain_data_list[0][0] ) ):

                        sem_above.append( ave_grain_data_list[0][1][k] + sem_grain_data_list[0][1][k] )
                        sem_below.append( ave_grain_data_list[0][1][k] + sem_grain_data_list[0][0][k] )

                    #Plot +/- SEM
                    plt.fill_between( ave_grain_data_list[0][0], sem_below, sem_above, facecolor = '#9999FF', interpolate = True )

                    plt.xlabel( 'cm^-1' )
                    plt.ylabel( 'Intensity (Normalized to {})'.format( normalized_to ) )

                    plt.ylim( 0., normalized_to*1.1 )

                    plt.xticks( np.arange( min( ave_grain_data_list[0][0] ), max( ave_grain_data_list[0][0] )+1, 500.) )
                    plt.yticks( np.arange( 0., normalized_to*1.1, step = 0.2*normalized_to ) ) 

                    data_line = mpatches.Patch(color='#000099', label='Averaged Data' )
                    sem_colors = mpatches.Patch(color='#9999FF', label='SEM Range')

                    plt.legend(handles=[data_line, sem_colors])

                    plt.title( "Averaged Data in Grain {}".format( ( i+1 ) ) )

                    plt.grid(True)

                    plt.savefig( os.path.join(ffn_out, "Grain {} - Averaged_RAMAN.png".format( (i+1) )) )
                        
                #If output type enabled, output data type
                if output_data:

                    #Open File
                    fle     = open( os.path.join(ffn_out, "Grain {} - Averaged_RAMAN{}".format( (i+1), exten )) , 'w' )

                    #Create Header
                    header  = "x" + delim + "y-Ave" + delim + "y-SEM\n"
                    fle.write( header )

                    #Cycle through 'x' values for SEM and
                    for l in range( len( sem_grain_data_list[0][0] ) ):
                        line = str( ave_grain_data_list[0][0][l] )
                        line += delim + str( ave_grain_data_list[0][1][l] )
                        line += delim + str( sem_grain_data_list[0][1][l] )
                        line += "\n"
                        fle.write( line )
                    
                    #Close File
                    fle.close()

                #Update Progress bar
                self.progress_var.set( prog_incr*(i+1) )
                self.mf.update_idletasks()

            self.progress_var.set( 0. )
            self.mf.update_idletasks()

    '''        
        #Process SPC files
        for fpath in flist:

            #Check if file ends with 'spc'
            if fpath.lower().endswith('spc'):

                #Get file name - the extension
                foutp = fpath[:-4] + exten
                try:
                    print(fpath, end=' ')

                    f = spc.File(fpath)
                    f.write_file(foutp, delimiter=delim)

                    print('Converted')
                except:
                    print('Error processing %s' % fpath)
            else:
                print('%s not spc file, skipping' % fpath)
        '''
    #Removes stray data points that aren't real measurements
    def sanitize_list(self, m_data, slope_threshold = 50. ):

        #Check data for large unwanted slopes
        
        for i in range( len( m_data[1] ) ):

            #If on first point, set the previous point
            if not i == 0:
                slope = abs( (m_data[1][i] - m_data[1][i-1] )/( m_data[0][i-1] - m_data[0][i] ) )

                #If the slope is too large, remove the point
                if slope >= slope_threshold:
                    
                    #If not the 2nd to last point
                    if not i == ( len( m_data[1] ) - 2 ):
                        
                        #Remove the large point and replace with an average
                        point_ave       = ( m_data[1][i-1] + m_data[1][i+1] )/2.
                        m_data[1][i]    = point_ave
        
        #Return the resultant list
        return m_data

    #Normalize given data to given max height
    def measurement_normalized( self, m_data, max_h, m_count):

        #Remove nasty stray data points
        clean_list = self.sanitize_list( m_data )

        #Find maximum value in list
        max_val = 0.

        #Cycle through 'y' values
        for i in range( len( clean_list[1] ) ):

            if  clean_list[1][i] > max_val:

                max_val =  clean_list[1][i]

        #Calculate Normalizing value
        norm_val = max_h/( max_val*m_count )

        #Normalize system to max_h/m_count
        for i in range( len( clean_list[1] ) ):
            
            clean_list[1][i] *= norm_val

        #Return Normalized data
        return clean_list

    #Average given data
    def grain_average( self, g_data, n_to ):
        #g_data = [ 
        #           [ [x,x,x,x], [y,y,y,y] ], 
        #           [ [x,x,x,x], [y,y,y,y] ], 
        #           [ [x,x,x,x], [y,y,y,y] ], 
        #          ... ]

        ave_grain_list = []

        #Cycle through each measurement on this grain
        for i in range( len( g_data ) ): #5

            #m_list = [ [x,x,x,x], [y,y,y,y] ]
            m_list = g_data[i]

            norm_m_list = self.measurement_normalized( m_list, n_to, len( g_data ) )  #( [[x],[y]], 1.0, 5 )

            #If first loop
            if i == 0:
                ave_grain_list.append( norm_m_list )

            #Else, add the list y-elements
            else:
                for j in range( len( ave_grain_list[0][1] ) ):
                    ave_grain_list[0][1][j] += norm_m_list[1][j]

        #Return Final List
        return ave_grain_list[0]

    #Standard Error of the Mean Calculation
    def SEM_calc(self, g_data, ave_g_data):

        p_sem = []
        m_sem = []

        sem_l = []

        #Cycle through each 'x' value
        for i in range( len( ave_g_data[0] ) ):

            #SEM at point i
            sem = 0.

            #Cycle through each 'y' value at the given 'x'
            count = len( g_data )           #5
            for j in range( count ):

                sem += math.pow( ( ave_g_data[1][i] - g_data[j][1][i] ) , 2. )

            sem = math.sqrt( sem )

            sem /= (count - 1.)*count

            p_sem.append( sem )
            m_sem.append( -1.*sem )
        
        sem_l.append( m_sem )
        sem_l.append( p_sem )

        return sem_l

    #Ask for directory. (Used in init)
    def ask_dir(self):
        self.folder.set( filedialog.askdirectory() )

    def ask_out_dir(self):
        self.output_folder.set( filedialog.askdirectory() )

    def output_message(self, title, message, button_txt, event):

        #Build Toplevel window and set its title
        self.sel_top = Toplevel(master = None, padx = 10, pady = 10)
        self.sel_top.title( title )

        self.sel_msg_txt = message

        #Show Message in window
        self.sel_msg = Message( self.sel_top, text = self.sel_msg_txt )
        self.sel_msg.pack()

        #Add button to window
        self.sel_button = Button( self.sel_top, text = button_txt, command = self.sel_top.destroy )
        self.sel_button.pack()

#Main call
if __name__ == "__main__":
    #Initiate TK
    root = Tk()

    #Initiate main class with root
    clss = AnalyzeSPC(root)

    #Set root as mainloop
    root.config( menu = clss.menubar )
    root.mainloop()
