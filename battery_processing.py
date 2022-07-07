# %%
class arbin_file:
    """methods for analyzing and plotting arbin battery cycling data"""
    ###### Open file ######    
    def __init__(self):
        """ Open xls arbin battery cycling files """
        
        # import modules
        import pyexcel as p
        import pandas as pd
        import matplotlib.pyplot as plt
        from matplotlib import colors
        from tkinter import Tk
        from tkinter.filedialog import askopenfilenames
        
        # select arbin file to load
        Tk().withdraw()
        ftuple=askopenfilenames()
        filenames=''.join(ftuple)
        xlsxfilenames=filenames+'x'
        
        # convert original .xls file to .xlsx file
            # pandas.read_excel() does not like .xls files
        p.save_book_as(file_name=filenames,
                   dest_file_name=xlsxfilenames)
    
        # read excel file and create a dictionary with sheetname:dataframe structure
        self.battery_dict=pd.read_excel(xlsxfilenames, sheet_name=None)
        self.sheetnames=list(self.battery_dict.keys()) 
        del self.sheetnames[0]
        
        # make dataframe listing channels, start times, and cell name
        info=self.battery_dict['Info'].iloc[:, [0,1,2]]
        info=info.dropna(axis=0)
        info=info.rename(columns=info.iloc[0])
        self.info=info.drop(info.index[0])
        
    
    def load_masses(self, masses, cells='all'):
        """ input cell loadings and calculate capacity (mAh/g) for each cell specified"""
         
        # map input parameters to call appropriate dataframes from battery_dict
        if cells!='all':
            cell_list_dict=dict(zip(cells, masses))
        
        if cells=='all':
            cells=[]
            for sheet in self.sheetnames:
                cell=sheet.split('-')[1]
                cells.append(cell)
            cells=list(dict.fromkeys(cells))
            cell_list_dict=dict(zip(cells, masses))
        
        # create dictionary of requested cells:dataframes
        sheets=[]
        for cell in cells:
            for sheet in self.battery_dict.keys():
                if cell in sheet:
                    sheets.append(sheet)
        
        # create dictionary of data frames for selected cells
        # calculate capacities in mAh/g
        # calculate dQ/dV
        sheet_dfs=[]
        for sheet in sheets:
            df=self.battery_dict[sheet]
            for cell in cells:
                if cell in sheet:
                    mass=cell_list_dict[cell]
                    df['Charge_Capacity(mAh/g)']=df['Charge_Capacity(Ah)']*1000/mass
                    df['Discharge_Capacity(mAh/g)']=df['Discharge_Capacity(Ah)']*1000/mass
                    if 'Statistics' in sheet:
                        df['Coulombic_Efficiency(%)']=(df['Charge_Capacity(Ah)']/df['Discharge_Capacity(Ah)'])*100       
            sheet_dfs.append(df)
        self.cells_dict=dict(zip(sheets, sheet_dfs))

                                      
        
    #def excel_igor():
        " Output excel file ready to import to igor using an igor module"
        
    def plot_voltage_profile(self, cells='all', cycles='all', axis_range=[0,500,1.5,5], legend_loc=0):
        """ plot the voltage profile for specified cells and cycle numbers"""
        
        # import modules
        import pandas as pd
        import matplotlib.pyplot as plt
        from matplotlib import colors
        
        
        # map cells input to cell_list
        if cells=='all':
            cell_list=[]
            for sheet in self.cells_dict.keys():
                cell=sheet.split('-')[1]
                cell_list.append(cell)
            cell_list=list(dict.fromkeys(cell_list))
        
        if cells!='all':
            cell_list=cells
        
        # create list of sheet names corresponding to selected cells
        sheets=[]
        for cell in cell_list:
            for sheet in self.cells_dict.keys():
                if 'Channel_1-'+cell in sheet:
                    sheets.append(sheet)
        
        # create dictionary of selected cells with corresponding list of cycles with dfs
        plot_dict={}
        for sheet in sheets:
            if cycles=='all':
                cycles=pd.Series(self.cells_dict[sheet]['Cycle_Index'], index=None)   
                cycles=cycles.drop_duplicates(keep='first').to_list()
            if cycles!='all':
                cycles=cycles
            cycle_dict={}
            for cycle in cycles:
                sheet_df=self.cells_dict[sheet]
                cycle_df=sheet_df.loc[sheet_df['Cycle_Index']==cycle]
                cycle_dict[cycle]=cycle_df
            plot_dict[sheet]=cycle_dict
        
        # create plot of specifed cells and specified cycles
        for sheet in plot_dict:
            cycle_list=[]
            for cycle in plot_dict[sheet].keys():
                if len(plot_dict[sheet][cycle])!=0:
                    cycle_list.append(cycle)
            norm = colors.Normalize(min(cycle_list), max(cycle_list))
            color_list=[]
            for cycle in cycle_list:
                color=plt.cm.rainbow_r(norm(cycle))
                color_list.append(color)
            color_map=dict(zip(cycle_list,color_list))
            self.figure,ax=plt.subplots(figsize=(8,6))
            for cycle in cycle_list:
                lb='Cycle'+str(cycle)
                df=plot_dict[sheet][cycle]
                dsch_voltage=df.loc[(df['Step_Index']==1) | (df['Step_Index']==2) | (df['Step_Index']==6)]['Voltage(V)']
                dsch_capacity=df.loc[(df['Step_Index']==1) | (df['Step_Index']==2) | (df['Step_Index']==6)]['Discharge_Capacity(mAh/g)']
                charge_voltage=df.loc[(df['Step_Index']==3) | (df['Step_Index']==4)]['Voltage(V)']
                charge_capacity=df.loc[(df['Step_Index']==3) | (df['Step_Index']==4)]['Charge_Capacity(mAh/g)']
                plt.plot(dsch_capacity, dsch_voltage,
                            label=lb, color=color_map[cycle], linewidth=2)
                plt.plot(charge_capacity,
                         charge_voltage,
                         color=color_map[cycle], linewidth=2)
                legend=plt.legend(fontsize=16, frameon=True, ncol=3, 
                           loc=legend_loc)
                legend.get_frame().set_linewidth(2)
                legend.get_frame().set_edgecolor('black')
                plt.xlabel('Capacity (mAh/g)', fontsize=20, labelpad=10)
                plt.ylabel('Potential (V vs Li+/Li)',fontsize=20, labelpad=10)
                plt.axis(axis_range)
                plt.setp(ax.spines.values(), linewidth=2)
                plt.tick_params(axis='both', direction='in', labelsize=18,
                                length=8, width=2)
            plt.show()
    
    # plot cycle life and coulombic efficiency
    def plot_cycle_life(self, cells='all', plot_type='both', 
                        y1_range=[0,50,400,550], y2_range=[0,50,90,105], 
                        decimals=0):
        """ plot cycle life and/or coulombic efficiency for selected cells"""
        
        # import modules
        import matplotlib.pyplot as plt
        from matplotlib import colors
        from matplotlib.ticker import FormatStrFormatter
        from matplotlib.text import OffsetFrom

        # map cells input to cell_list
        if cells=='all':
            cell_list=[]
            for sheet in self.cells_dict.keys():
                cell=sheet.split('-')[1]
                cell_list.append(cell)
            cell_list=list(dict.fromkeys(cell_list))

        if cells!='all':
            cell_list=cells
        
        # map decimals input to variable
        decimals='%.'+str(decimals)+'f'

        # create list of sheet names corresponding to selected cells
        sheets=[]
        for cell in cell_list:
            for sheet in self.cells_dict.keys():
                if 'Statistics_1-'+cell in sheet:
                    sheets.append(sheet)
            sheets=list(dict.fromkeys(sheets))

        # plot cycle life and/or coulombic efficiency for each cell selected
        for sheet in sheets:
            cycles=self.cells_dict[sheet]['Cycle_Index']
            capacity=self.cells_dict[sheet]['Discharge_Capacity(mAh/g)']
            coulomb_eff=self.cells_dict[sheet]['Coulombic_Efficiency(%)']
            figure,ax1=plt.subplots(figsize=(8,6))

            if plot_type=='capacity':
                plt.plot(cycles, capacity, 
                         color='red', marker='o', linewidth=0)
                plt.ylabel('Capacity (mAh/g)',fontsize=20, labelpad=10)

            if plot_type=='coulombic':
                plt.plot(cycles, coulomb_eff, 
                         color='red', marker='o', linewidth=0)
                plt.ylabel('Coulombic Efficiency (%)',fontsize=20, labelpad=10)

            if plot_type=='both':
                ax1.plot(cycles, capacity, label='Capacity', 
                         color='red', marker='o', linewidth=0)
                ax1.set_ylabel('Capacity (mAh/g)', fontsize=20, labelpad=10, color='red')
                ax2=ax1.twinx()
                ax2.plot(cycles, coulomb_eff,
                         label='Coulombic Efficiency', color='blue', 
                         marker='o', linewidth=0)
                ax2.set_ylabel('Coulombic Efficiency (%)', fontsize=20, labelpad=10, color='blue')
                ax2.tick_params(axis='both', direction='in', labelsize=18,
                            length=8, width=2)
                ax2.ticklabel_format(useOffset=False) 
                ax2.yaxis.set_major_formatter(FormatStrFormatter(decimals))
                ax2.axis(y2_range)

            ax1.set_xlabel('Cycle #', fontsize=20, labelpad=10)
            plt.setp(ax1.spines.values(), linewidth=2)
            ax1.tick_params(axis='both', direction='in', labelsize=18,
                            length=8, width=2)
            ax1.ticklabel_format(useOffset=False) 
            ax1.yaxis.set_major_formatter(FormatStrFormatter(decimals))
            ax1.axis(y1_range)

            plt.show()       
    
# %%
class neware_file():
    """methods for analyzing and plotting neware battery cycling data"""
    ###### Open file ######    
    def __init__(self):
        """ Open xls neware battery cycling files """
        
        # import modules
        import pyexcel as p
        import pandas as pd
        from tkinter import Tk
        from tkinter.filedialog import askopenfilenames
        
        # select arbin file to load
        Tk().withdraw()
        ftuple=askopenfilenames()
        filenames=''.join(ftuple)
        xlsxfilenames=filenames+'x'
        
        # read excel file and create a dictionary with sheetname:dataframe structure
        self.battery_dict=pd.read_excel(xlsxfilenames, sheet_name=None)
        self.sheetnames=list(self.battery_dict.keys())
        del self.sheetnames[0]
        
        # make dataframe listing channels, start times, and cell name
        info=self.battery_dict['Info'].iloc[:, [0,1,2,5]]
        file_name=info.iat[0,0]
        info['File_Name']=file_name
        info=info.dropna(axis=0)
        info=info.rename(columns=info.iloc[0])
        self.info=info.drop(info.index[0])

    def load_masses(self, masses, cells='all'):
        """ input cell loadings and calculate capacity (mAh/g) for each cell specified"""

        # import modules
        import pandas as pd
        from numpy import diff

        # map input parameters to call appropriate dataframes from battery_dict
        if cells!='all':
            cell_list_dict=dict(zip(cells, masses))

        if cells=='all':
            cells=[]
            for sheet in self.sheetnames:
                strings=sheet.split('_')
                cell=strings[2]+'_'+strings[3]
                cells.append(cell)
            cells=list(dict.fromkeys(cells))
            cell_list_dict=dict(zip(cells, masses))

        # create dictionary of requested cells:dataframes
        sheets=[]
        for cell in cells:
            for sheet in self.battery_dict.keys():
                if cell in sheet:
                    sheets.append(sheet)

        # create dictionary of data frames for selected cells
        # calculate capacities in mAh/g
        # calculate dQ/dV
        sheet_dfs=[]
        for sheet in sheets:
            df=self.battery_dict[sheet]
            for cell in cells:
                if cell in sheet:
                    mass=cell_list_dict[cell]
                    if 'Cycle' in sheet:
                        df['Charge_Capacity(mAh/g)']=df['Capacity of charge(mAh)']/mass
                        df['Discharge_Capacity(mAh/g)']=df['Capacity of discharge(mAh)']/mass
                        df['Coulombic_Efficiency(%)']=(df['Capacity of charge(mAh)']/df['Capacity of discharge(mAh)'])*100 
                    if 'Detail' in sheet:
                        df['Capacity(mAh/g)']=df['CapaCity(mAh)']/mass
                        dqdv_starter=[0]
                        df2=pd.DataFrame()
                        df1=pd.DataFrame(dqdv_starter, columns=['dQ/dV'])
                        df2['dQ/dV']=diff(df['CapaCity(mAh)'])/diff(df['Voltage(V)'])
                        dqdv=pd.concat([df1,df2])
                        df=df.join(dqdv)        
            sheet_dfs.append(df)
        self.cells_dict=dict(zip(sheets, sheet_dfs))
        
    #def excel_igor():
        " Output excel file ready to import to igor using an igor module"
        
    def plot_voltage_profile(self, cells='all', cycles='all', axis_range=[0,500,1.5,5], legend_loc=0):
        """ plot the voltage profile for specified cells and cycle numbers"""
        
        # import modules
        import pandas as pd
        import matplotlib.pyplot as plt
        from matplotlib import colors
        
        
        # map cells input to cell_list
        if cells=='all':
            cell_list=[]
            for sheet in self.sheetnames:
                strings=sheet.split('_')
                cell=strings[2]+'_'+strings[3]
                cell_list.append(cell)
                cell_list.append(cell)
                cell_list=list(dict.fromkeys(cell_list))

        if cells!='all':
            cell_list=cells

        # create list of sheet names corresponding to selected cells
        sheets=[]
        for cell in cell_list:
            for sheet in self.cells_dict.keys():
                if 'Detail_114_'+cell in sheet:
                    sheets.append(sheet)

        # create dictionary of selected cells with corresponding list of cycles with dfs
        plot_dict={}
        for sheet in sheets:
            if cycles=='all':
                cycles=pd.Series(self.cells_dict[sheet]['Cycle_Index'], index=None)   
                cycles=cycles.drop_duplicates(keep='first').to_list()
            if cycles!='all':
                cycles=cycles
            cycle_dict={}
            for cycle in cycles:
                sheet_df=self.cells_dict[sheet]
                cycle_df=sheet_df.loc[sheet_df['Cycle']==cycle]
                cycle_dict[cycle]=cycle_df
            plot_dict[sheet]=cycle_dict
            
        # create plot of specifed cells and specified cycles
        for sheet in plot_dict:
            cycle_list=[]
            for cycle in plot_dict[sheet].keys():
                if len(plot_dict[sheet][cycle])!=0:
                    cycle_list.append(cycle)
            norm = colors.Normalize(min(cycle_list), max(cycle_list))
            color_list=[]
            for cycle in cycle_list:
                color=plt.cm.rainbow_r(norm(cycle))
                color_list.append(color)
            color_map=dict(zip(cycle_list,color_list))
            self.figure,ax=plt.subplots(figsize=(8,6))
            for cycle in cycle_list:
                lb='Cycle'+str(cycle)
                df=plot_dict[sheet][cycle]
                dsch_voltage=df.loc[(df['Status']=='CC_DChg')]['Voltage(V)']
                dsch_capacity=df.loc[(df['Status']=='CC_DChg')]['Capacity(mAh/g)']
                charge_voltage=df.loc[(df['Status']=='CC_Chg')]['Voltage(V)']
                charge_capacity=df.loc[(df['Status']=='CC_Chg')]['Capacity(mAh/g)']
                plt.plot(dsch_capacity, dsch_voltage,
                            label=lb, color=color_map[cycle], linewidth=2)
                plt.plot(charge_capacity,
                         charge_voltage,
                         color=color_map[cycle], linewidth=2)
                legend=plt.legend(fontsize=16, frameon=True, ncol=3, 
                           loc=legend_loc)
                legend.get_frame().set_linewidth(2)
                legend.get_frame().set_edgecolor('black')
                plt.xlabel('Capacity (mAh/g)', fontsize=20, labelpad=10)
                plt.ylabel('Potential (V vs Li+/Li)',fontsize=20, labelpad=10)
                plt.axis(axis_range)
                plt.setp(ax.spines.values(), linewidth=2)
                plt.tick_params(axis='both', direction='in', labelsize=18,
                                length=8, width=2)
            plt.show()
            
    # plot cycle life and coulombic efficiency
    def plot_cycle_life(self, cells='all', plot_type='both', 
                        y1_range=[0,50,400,550], y2_range=[0,50,90,105], 
                        decimals=0):
        """ plot cycle life and/or coulombic efficiency for selected cells"""
        
        # import modules
        import matplotlib.pyplot as plt
        from matplotlib import colors
        from matplotlib.ticker import FormatStrFormatter
        from matplotlib.text import OffsetFrom  
        
        # map cells input to cell_list
        if cells=='all':
            cell_list=[]
            for sheet in self.sheetnames:
                strings=sheet.split('_')
                cell=strings[2]+'_'+strings[3]
                cell_list.append(cell)
                cell_list.append(cell)
                cell_list=list(dict.fromkeys(cell_list))

        if cells!='all':
            cell_list=cells
        
        # map decimals input to variable
        decimals='%.'+str(decimals)+'f'

        # create list of sheet names corresponding to selected cells
        sheets=[]
        for cell in cell_list:
            for sheet in self.cells_dict.keys():
                if 'Cycle_114_'+cell in sheet:
                    sheets.append(sheet)
            sheets=list(dict.fromkeys(sheets))    
        
        # plot cycle life and/or coulombic efficiency for each cell selected
        for sheet in sheets:
            cycles=self.cells_dict[sheet]['ToTal of Cycle']
            capacity=self.cells_dict[sheet]['Discharge_Capacity(mAh/g)']
            coulomb_eff=self.cells_dict[sheet]['Coulombic_Efficiency(%)']
            figure,ax1=plt.subplots(figsize=(8,6))

            if plot_type=='capacity':
                plt.plot(cycles, capacity, 
                         color='red', marker='o', linewidth=0)
                plt.ylabel('Capacity (mAh/g)',fontsize=20, labelpad=10)

            if plot_type=='coulombic':
                plt.plot(cycles, coulomb_eff, 
                         color='red', marker='o', linewidth=0)
                plt.ylabel('Coulombic Efficiency (%)',fontsize=20, labelpad=10)

            if plot_type=='both':
                ax1.plot(cycles, capacity, label='Capacity', 
                         color='red', marker='o', linewidth=0)
                ax1.set_ylabel('Capacity (mAh/g)', fontsize=20, labelpad=10, color='red')
                ax2=ax1.twinx()
                ax2.plot(cycles, coulomb_eff,
                         label='Coulombic Efficiency', color='blue', 
                         marker='o', linewidth=0)
                ax2.set_ylabel('Coulombic Efficiency (%)', fontsize=20, labelpad=10, color='blue')
                ax2.tick_params(axis='both', direction='in', labelsize=18,
                            length=8, width=2)
                ax2.ticklabel_format(useOffset=False) 
                ax2.yaxis.set_major_formatter(FormatStrFormatter(decimals))
                ax2.axis(y2_range)

            ax1.set_xlabel('Cycle #', fontsize=20, labelpad=10)
            plt.setp(ax1.spines.values(), linewidth=2)
            ax1.tick_params(axis='both', direction='in', labelsize=18,
                            length=8, width=2)
            ax1.ticklabel_format(useOffset=False) 
            ax1.yaxis.set_major_formatter(FormatStrFormatter(decimals))
            ax1.axis(y1_range)

            plt.show() 
# %%
