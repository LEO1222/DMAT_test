#!/usr/bin/env python3
import mhi.pscad
import logging
import mhi.pscad.handler
import os, openpyxl
import pandas as pd


class BuildEventHandler(mhi.pscad.handler.BuildEvent):

    def __init__(self):
        super().__init__()
        self._start = {}

    def _build_event(self, phase, status, project, elapsed, **kwargs):

        key = (project, phase)
        if status == 'BEGIN':
            self._start[key] = elapsed
        else:
            sec = elapsed - self._start[key]
            name = project if project else '[All]'
            LOG.info("%s %s: %.3f sec", name, phase, sec)


# Log 'INFO' messages & above.  Include level & module name.
logging.basicConfig(level=logging.INFO,
                    format="%(levelname)-8s %(name)-26s %(message)s")

# Ignore INFO msgs from automation (eg, mhi.pscad, mhi.pscad.pscad, ...)
logging.getLogger('mhi.pscad').setLevel(logging.WARNING)

LOG = logging.getLogger('main')

versions = mhi.pscad.versions()
LOG.info("PSCAD Versions: %s", versions)

# Skip any 'Alpha' versions, if other choices exist
vers = [(ver, x64) for ver, x64 in versions if ver != 'Alpha']
if len(vers) > 0:
    versions = vers

# Skip any 'Beta' versions, if other choices exist
vers = [(ver, x64) for ver, x64 in versions if ver != 'Beta']
if len(vers) > 0:
    versions = vers

# Skip any 32-bit versions, if other choices exist
vers = [(ver, x64) for ver, x64 in versions if x64]
if len(vers) > 0:
    versions = vers

LOG.info("   After filtering: %s", versions)

# Of any remaining versions, choose the "lexically largest" one.
version, x64 = sorted(versions)[-1]
LOG.info("   Selected PSCAD version: %s %d-bit", version, 64 if x64 else 32)

# # Get all installed FORTRAN compiler versions
# fortrans = mhi.pscad.fortran_versions()
# LOG.info("FORTRAN Versions: %s", fortrans)

# # Skip 'GFortran' compilers, if other choices exist
# vers = [ver for ver in fortrans if 'GFortran' not in ver]
# if len(vers) > 0:
#     fortrans = vers

# LOG.info("   After filtering: %s", fortrans)

# # Order the remaining compilers, choose the last one (highest revision)
# fortran = sorted(fortrans)[-1]
# LOG.info("   Selected FORTRAN version: %s", fortran)

# Get all installed Matlab versions
matlabs = mhi.pscad.matlab_versions()
LOG.info("Matlab Versions: %s", matlabs)

# Get the highest installed version of Matlab:
matlab = sorted(matlabs)[-1] if matlabs else ''
LOG.info("   Selected Matlab version: %s", matlab)

# Launch PSCAD
LOG.info("Launching: %s  FORTRAN=%r   Matlab=%r",
         version, 'Not Available', matlab)
pscad = mhi.pscad.launch(minimize=False, version=version, x64=x64)

if pscad:

    try:

        # Load only the pscx project file
        pscad.load(r"C:\Users\Niu2021\Desktop\integration\tests_integration.pscx")

        # Get the list of simulation sets
        sim_sets = pscad.simulation_sets()
        if len(sim_sets) > 0:
            LOG.info("Simulation sets: %s", sim_sets)

            # For each simulation set ...
            for sim_set_name in sim_sets:
                # ... run it
                LOG.info("Running simulation set '%s'", sim_set_name)
                sim_set = pscad.simulation_set(sim_set_name)
                sim_set.run()
                LOG.info("Simulation set '%s' complete", sim_set_name)
        else:
            # Run project
            
            FaultTest = pscad.project("tests_integration")


            ############################120_Large_Disturbance_Test#########################
            ############################# Determine excel file path ##################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'large_disturbance120.xlsx')
            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame
            # DataFrame -----> Python List
            data_value = df.values.tolist()
            ############################################################################################
            
            #generate variable list
            list_Duration = (data_value[0])[1:] #variable 1
            list_FaultType = (data_value[1])[1:]
            list_P = (data_value[6])[1:]
            list_Q = (data_value[7])[1:]
            list_Rs = (data_value[4])[1:]
            list_Xs = (data_value[5])[1:]
            list_Rf = (data_value[2])[1:]
            list_Xf = (data_value[3])[1:] #variable 8


            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas
            canvas1 = FaultTest.canvas("Grid_Side_Ctrl") # get the controller of grid side controller canvas
            
            # Use canvas controller to find components by name
            Duration = canvas.find("master:const", "Duration_Setting")
            FaultType = canvas.find("master:const", "FaultType_Setting")
            Q = canvas1.find("master:const", "Qref_DMAT") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")
            Rf = canvas.find("master:const", "Rfault")
            Xf = canvas.find("master:const", "Xfault")


            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")

            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Disabled"
            figure9.state = "Disabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Disabled"
            table13minus40.state = "Disabled"
            table1360.state = "Disabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Enabled"

            # Run each case 
            for index in range(len(list_Rs)):                
                # Change variables each cycle
                Duration.parameters(Name="Duration_Setting", Value=list_Duration[index])
                FaultType.parameters(Name="FaultType_Setting", Value=list_FaultType[index])
                Q.parameters(Name="Qref_DMAT", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])
                Rf.parameters(Name="Rfault", Value=list_Rf[index])
                Xf.parameters(Name="Xfault", Value=list_Xf[index])

                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"large_disturbance120_test{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)

            ############################## Muyuan_summerTerm #######################
            ############################## Determine excel file path ###############################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'muyuan_part.xlsx')
            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame
            ## DataFrame -----> Python List
            data_value = df.values.tolist()
            
            #generate variable list
            lst_switch1 = (data_value[0])[1:] #variable 1
            lst_switch2 = (data_value[1])[1:]
            lst_switch3 = (data_value[2])[1:]
            lst_switch4 = (data_value[3])[1:]
            lst_switch5 = (data_value[4])[1:]
            lst_switch6 = (data_value[5])[1:]
            lst_Rgrid = (data_value[6])[1:]
            lst_Xgrid = (data_value[7])[1:] #variable 8

            # Select the specific component
            canvas0 = FaultTest.canvas("Main") # get the controller of main canvas
            canvas1 = FaultTest.canvas("Grid_Side_Ctrl") # get the controller of grid side controller canvas
            canvas2 = FaultTest.canvas("Machin_Side_Ctrl") # get the controller of Machine side controller canvas
            canvas3 = FaultTest.canvas("WindTurbine_Mechanical") #get the controller of WindTurbine controller canvas

            # Use canvas controller to find components by name
            switch1 = canvas1.find("master:const", "switch1")
            switch2 = canvas1.find("master:const", "switch2")
            switch3 = canvas2.find("master:const", "switch3") #in grid-side controller
            switch4 = canvas2.find("master:const", "switch4")
            switch5 = canvas3.find("master:const", "switch5")
            switch6 = canvas3.find("master:const", "switch6")
            Rgrid = canvas0.find("master:const", "Rgrid")
            Xgrid = canvas0.find("master:const", "Xgrid")
            
            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")


            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Disabled"
            figure9.state = "Disabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Disabled"
            table13minus40.state = "Disabled"
            table1360.state = "Disabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Enabled"
            NonMuyuan.state = "Disabled"
            large1to120.state = "Disabled"
            
            # Run each case 
            for index in range(len(lst_switch1)):
                # Change variables each cycle
                switch1.parameters(Name="switch1", Value=lst_switch1[index])
                switch2.parameters(Name="switch2", Value=lst_switch2[index])
                switch3.parameters(Name="switch3", Value=lst_switch3[index])
                switch4.parameters(Name="switch4", Value=lst_switch4[index])
                switch5.parameters(Name="switch5", Value=lst_switch5[index])
                switch6.parameters(Name="switch6", Value=lst_switch6[index])
                Rgrid.parameters(Name="Rgrid", Value=lst_Rgrid[index])
                Xgrid.parameters(Name="Xgrid", Value=lst_Xgrid[index])
                
                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"summer_term_muyuan_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)
            
            ######################### Yuxiang 206FRT #################################################
            ############################# Determine excel file path ##################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'yuxiang_206_225FRT.xlsx')
            print('The xlsFile path is:',xlsPath)
            print('-'*60)

            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame
            # DataFrame -----> Python List
            data_value = df.values.tolist()
            ############################################################################################
            #generate variable list
            list_Duration = (data_value[0])[1:] #variable 1
            list_FaultType = (data_value[1])[1:]
            list_P = (data_value[6])[1:]
            list_Q = (data_value[7])[1:]
            list_Rs = (data_value[4])[1:]
            list_Xs = (data_value[5])[1:]
            list_Rf = (data_value[2])[1:]
            list_Xf = (data_value[3])[1:] #variable 8


            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas
            canvas1 = FaultTest.canvas("Grid_Side_Ctrl") # get the controller of grid side controller canvas
            
            # Use canvas controller to find components by name
            Duration = canvas.find("master:const", "Duration_Setting")
            FaultType = canvas.find("master:const", "FaultType_Setting")
            Q = canvas1.find("master:const", "Qref_DMAT") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")
            Rf = canvas.find("master:const", "Rfault")
            Xf = canvas.find("master:const", "Xfault")


            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")

            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Disabled"
            figure9.state = "Disabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Disabled"
            table13minus40.state = "Disabled"
            table1360.state = "Disabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Enabled"

            # Run each case 
            for index in range(len(list_Rs)):                
                # Change variables each cycle
                Duration.parameters(Name="Duration_Setting", Value=list_Duration[index])
                FaultType.parameters(Name="FaultType_Setting", Value=list_FaultType[index])
                Q.parameters(Name="Qref_DMAT", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])
                Rf.parameters(Name="Rfault", Value=list_Rf[index])
                Xf.parameters(Name="Xfault", Value=list_Xf[index])

                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"yuxiang_206FRT_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)

            ########################################## Yuxiang Part ###############################################################
            ######################################### Determine excel file path ##############################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'yuxiang_part.xlsx')

            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame
            # DataFrame -----> Python List
            data_value = df.values.tolist()
            ####################################################################################################################
            #generate variable list
            list_Duration = (data_value[0])[1:]
            list_FaultType = (data_value[1])[1:]
            list_P = (data_value[6])[1:]
            list_Q = (data_value[7])[1:]
            list_Rs = (data_value[4])[1:]
            list_Xs = (data_value[5])[1:]
            list_Rf = (data_value[2])[1:]
            list_Xf = (data_value[3])[1:] 
            list_Rs_post=(data_value[8])[1:]
            list_Xs_post = (data_value[9])[1:]

            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas
            
            # Use canvas controller to find components by name
            Duration = canvas.find("master:const", "Duration_Setting")
            FaultType = canvas.find("master:const", "FaultType_Setting")
            Q = canvas.find("master:const", "Q_Setting") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")
            Rf = canvas.find("master:const", "Rfault")
            Xf = canvas.find("master:const", "Xfault")
            Rs_post = canvas.find("master:const", "Rgrid_post")
            Xs_post = canvas.find("master:const", "Xgrid_post")
            
        
            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")

            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Disabled"
            figure9.state = "Disabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Disabled"
            table13minus40.state = "Disabled"
            table1360.state = "Disabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Enabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Disabled"
            
            # Run each case
            for index in range(len(list_Rs)):                
                # Change variables each cycle
                Duration.parameters(Name="Duration_Setting", Value=list_Duration[index])
                FaultType.parameters(Name="FaultType_Setting", Value=list_FaultType[index])
                Q.parameters(Name="Q_Setting", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])
                Rf.parameters(Name="Rfault", Value=list_Rf[index])
                Xf.parameters(Name="Xfault", Value=list_Xf[index])
                Rs_post.parameters(Name="Rgrid_post", Value=list_Rs_post[index])
                Xs_post.parameters(Name="Xgrid_post", Value=list_Xs_post[index])

                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"yuxiang_summer_term_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)

            ################################## Ziheng 170Blue #######################################
            ######################################## Determine excel file path ########################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'test170-173.xlsx')

            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame
            # DataFrame -----> Python List
            data_value = df.values.tolist()
            ########################################################################################################################
            
            #generate variable list
            list_P = (data_value[2])[1:]
            list_Q = (data_value[3])[1:]
            list_Rs = (data_value[0])[1:]
            list_Xs = (data_value[1])[1:]

            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas
            
            # Use canvas controller to find components by name
            Q = canvas.find("master:const", "Q_Setting") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")
            
            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")

            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Disabled"
            figure9.state = "Disabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Enabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Disabled"
            table13minus40.state = "Disabled"
            table1360.state = "Disabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Disabled"

            # Run each case (168)
            for index in range(len(list_Rs)):

                # Change variables each cycle
                Q.parameters(Name="Q_Setting", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])


                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"ziheng_170blue_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output) 

            ######################## Ziheng 170 grey ############################################    
            ############################################### Determine excel file path ###############################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'test170-173.xlsx')

            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame

            # DataFrame -----> Python List
            data_value = df.values.tolist()
            ##############################################################################################################################
            
            #generate variable list
            list_P = (data_value[2])[1:]
            list_Q = (data_value[3])[1:]
            list_Rs = (data_value[0])[1:]
            list_Xs = (data_value[1])[1:]


            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas
            
            # Use canvas controller to find components by name
            Q = canvas.find("master:const", "Q_Setting") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")
            
            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")

            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Disabled"
            figure9.state = "Disabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Enabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Disabled"
            table13minus40.state = "Disabled"
            table1360.state = "Disabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Disabled"

            # Run each case (168)
            for index in range(len(list_Rs)):

                # Change variables each cycle
                Q.parameters(Name="Q_Setting", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])

                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"t{index+1}.out")
                FaultTest.run()


                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"ziheng_170grey_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)
            
            ################################# Ziheng 170 orange ############################################
            ############################################# Determine excel file path #############################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'test170-173.xlsx')
            print('The xlsFile path is:',xlsPath)
            print('-'*60)

            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame
            # DataFrame -----> Python List
            data_value = df.values.tolist()
            #####################################################################################################################
            #generate variable list
            list_P = (data_value[2])[1:]
            list_Q = (data_value[3])[1:]
            list_Rs = (data_value[0])[1:]
            list_Xs = (data_value[1])[1:]


            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas
            
            # Use canvas controller to find components by name
            Q = canvas.find("master:const", "Q_Setting") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")
            
            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")

            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Disabled"
            figure9.state = "Disabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Enabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Disabled"
            table13minus40.state = "Disabled"
            table1360.state = "Disabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Disabled"

            # Run each case 
            for index in range(len(list_Rs)):
                
                # Change variables each cycle
                Q.parameters(Name="Q_Setting", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])


                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"ziheng_170orange_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)
            
            ##################################### Ziheng 170 yellow ###############################################################
            ######################################### Determine excel file path #############################################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'test170-173.xlsx')

            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame

            # DataFrame -----> Python List
            data_value = df.values.tolist()
            #################################################################################################################################
            
            #generate variable list
            list_P = (data_value[2])[1:]
            list_Q = (data_value[3])[1:]
            list_Rs = (data_value[0])[1:]
            list_Xs = (data_value[1])[1:]

            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas
            
            # Use canvas controller to find components by name
            Q = canvas.find("master:const", "Q_Setting") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")
            
            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")

            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Disabled"
            figure9.state = "Disabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Enabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Disabled"
            table13minus40.state = "Disabled"
            table1360.state = "Disabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Disabled"

            # Run each case (168)
            for index in range(len(list_Rs)):
        
                # Change variables each cycle
                Q.parameters(Name="Q_Setting", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])


                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"ziheng_yellow_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)
            
            ############################################ Ziheng 174 blue ############################################
            ################################################# Determine excel file path #################################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'test174-177.xlsx')

            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame

            # DataFrame -----> Python List
            data_value = df.values.tolist()
            #################################################################################################################################
            #generate variable list
            list_P = (data_value[2])[1:]
            list_Q = (data_value[3])[1:]
            list_Rs = (data_value[0])[1:]
            list_Xs = (data_value[1])[1:]


            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas
            
            # Use canvas controller to find components by name
            Q = canvas.find("master:const", "Q_Setting") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")

            
            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")

            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Disabled"
            figure9.state = "Disabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Enabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Disabled"
            table13minus40.state = "Disabled"
            table1360.state = "Disabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Disabled"

            # Run each case (168)
            for index in range(len(list_Rs)):

                # Change variables each cycle
                Q.parameters(Name="Q_Setting", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])


                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"ziheng_174blue_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)


            ########################################## Ziheng 174 yellow ############################################
            ################################################ Determine excel file path ################################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'test174-177.xlsx')

            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame

            # DataFrame -----> Python List
            data_value = df.values.tolist()
            ##############################################################################################################################
            
            #generate variable list
            list_P = (data_value[2])[1:]
            list_Q = (data_value[3])[1:]
            list_Rs = (data_value[0])[1:]
            list_Xs = (data_value[1])[1:]

            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas
            
            # Use canvas controller to find components by name
            Q = canvas.find("master:const", "Q_Setting") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")

            
            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")
            large1to120.state = "Disabled"

            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Disabled"
            figure9.state = "Disabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Enabled"
            table1340.state = "Disabled"
            table13minus40.state = "Disabled"
            table1360.state = "Disabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Disabled"       
            
            # Run each case (168)
            for index in range(len(list_Rs)):

                # Change variables each cycle
                Q.parameters(Name="Q_Setting", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])

                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"ziheng_174yellow_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)


            ##################################### Ziheng 186 Blue #######################################
            ######################################## Determine excel file path ####################################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'test186-189.xlsx')

            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame

            # DataFrame -----> Python List
            data_value = df.values.tolist()
            ########################################################################################################################
            
            #generate variable list
            list_P = (data_value[2])[1:]
            list_Q = (data_value[3])[1:]
            list_Rs = (data_value[0])[1:]
            list_Xs = (data_value[1])[1:]

            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas
            
            # Use canvas controller to find components by name
            Q = canvas.find("master:const", "Q_Setting") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")

            
            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")

            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Disabled"
            figure9.state = "Disabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Enabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Disabled"
            table13minus40.state = "Disabled"
            table1360.state = "Disabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Disabled"

            # Run each case (168)
            for index in range(len(list_Rs)):
                
                # Change variables each cycle
                Q.parameters(Name="Q_Setting", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])


                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"ziheng_186blue_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)

            ######################################## Ziheng 186 green #######################################
            ################################################## Determine excel file path ################################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'test186-189.xlsx')

            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame

            # DataFrame -----> Python List
            data_value = df.values.tolist()
            ###############################################################################################################################
            #generate variable list
            list_P = (data_value[2])[1:]
            list_Q = (data_value[3])[1:]
            list_Rs = (data_value[0])[1:]
            list_Xs = (data_value[1])[1:]

            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas
            
            # Use canvas controller to find components by name
            Q = canvas.find("master:const", "Q_Setting") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")

            
            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")

            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Disabled"
            figure9.state = "Disabled"
            figure10green.state = "Enabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Disabled"
            table13minus40.state = "Disabled"
            table1360.state = "Disabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Disabled"

            # Run each case (168)
            for index in range(len(list_Rs)):
                
                # Change variables each cycle
                Q.parameters(Name="Q_Setting", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])


                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"ziheng_186green_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)                

            ########################################## Ziheng 186 red ############################################
            ################################################ Determine excel file path ################################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'test186-189.xlsx')

            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame

            # DataFrame -----> Python List
            data_value = df.values.tolist()
            #############################################################################################################################
            #generate variable list
            list_P = (data_value[2])[1:]
            list_Q = (data_value[3])[1:]
            list_Rs = (data_value[0])[1:]
            list_Xs = (data_value[1])[1:]


            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas

            
            # Use canvas controller to find components by name
            Q = canvas.find("master:const", "Q_Setting") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")

            
            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")

            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Disabled"
            figure9.state = "Disabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Enabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Disabled"
            table13minus40.state = "Disabled"
            table1360.state = "Disabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Disabled"

            # Run each case (168)
            for index in range(len(list_Rs)):
                
                # Change variables each cycle
                Q.parameters(Name="Q_Setting", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])
                # Rf.parameters(Name="Rfault", Value=list_Rf[index])
                # Xf.parameters(Name="Xfault", Value=list_Xf[index])


                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"ziheng_186red_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)


            ############################################ Ziheng 190-192 #######################################
            ################################################## Determine excel file path ################################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'test190-192.xlsx')

            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame

            # DataFrame -----> Python List
            data_value = df.values.tolist()
            ###############################################################################################################################
            #generate variable list
            list_P = (data_value[2])[1:]
            list_Q = (data_value[3])[1:]
            list_Rs = (data_value[0])[1:]
            list_Xs = (data_value[1])[1:]
            list_Frequency = (data_value[4])[1:]


            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas
            
            # Use canvas controller to find components by name
            Q = canvas.find("master:const", "Q_Setting") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")
            Frequency= canvas.find("master:const", "Oscillatory Frequ")

            
            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")

            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Disabled"
            figure9.state = "Disabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Enabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Disabled"
            table13minus40.state = "Disabled"
            table1360.state = "Disabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Disabled"

            # Run each case (168)
            for index in range(len(list_Rs)):
                
                # Change variables each cycle
                Q.parameters(Name="Q_Setting", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])
                Frequency.parameters(Name="Oscillatory Frequ", Value=list_Frequency[index])

                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"ziheng_190_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)

            ############################################ 193_198_40 #################################################################
            ############################################## Determine excel file path ################################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'test193-198.xlsx')
            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame
            # DataFrame -----> Python List
            data_value = df.values.tolist()
            ###########################################################################################################################
            #generate variable list
            list_P = (data_value[2])[1:]
            list_Q = (data_value[3])[1:]
            list_Rs = (data_value[0])[1:]
            list_Xs = (data_value[1])[1:]


            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas
            canvas1 = FaultTest.canvas("Grid_Side_Ctrl") # get the controller of grid side controller canvas
            
            # Use canvas controller to find components by name
            Q = canvas.find("master:const", "Q_Setting") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")

            
            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")

            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Disabled"
            figure9.state = "Disabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Enabled"
            table13minus40.state = "Disabled"
            table1360.state = "Disabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Disabled"

            # Run each case (168)
            for index in range(len(list_Rs)):
                
                # Change variables each cycle
                Q.parameters(Name="Q_Setting", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])


                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"ziheng_193_+40_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)

            ################################## Ziheng 193_198_60 ######################################
            ############################################## Determine excel file path #####################################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'test193-198.xlsx')

            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame

            # DataFrame -----> Python List
            data_value = df.values.tolist()
            #################################################################################################################################
            #generate variable list
            list_P = (data_value[2])[1:]
            list_Q = (data_value[3])[1:]
            list_Rs = (data_value[0])[1:]
            list_Xs = (data_value[1])[1:]


            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas

            
            # Use canvas controller to find components by name
            Q = canvas.find("master:const", "Q_Setting") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")

            
            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")

            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Disabled"
            figure9.state = "Disabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Disabled"
            table13minus40.state = "Disabled"
            table1360.state = "Enabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Disabled"

            # Run each case (168)
            for index in range(len(list_Rs)):

                # Change variables each cycle
                Q.parameters(Name="Q_Setting", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])


                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"ziheng_193_+60_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)

            ############################ Ziheng 193 198_-40 #################################
            ############################################################ Determine excel file path##############################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'test193-198.xlsx')

            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame

            # DataFrame -----> Python List
            data_value = df.values.tolist()
            #####################################################################################################################################
            #generate variable list
            list_P = (data_value[2])[1:]
            list_Q = (data_value[3])[1:]
            list_Rs = (data_value[0])[1:]
            list_Xs = (data_value[1])[1:]


            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas

            
            # Use canvas controller to find components by name
            Q = canvas.find("master:const", "Q_Setting") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")
            
            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")

            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Disabled"
            figure9.state = "Disabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Disabled"
            table13minus40.state = "Enabled"
            table1360.state = "Disabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Disabled"

            # Run each case (168)
            for index in range(len(list_Rs)):
                
                # Change variables each cycle
                Q.parameters(Name="Q_Setting", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])


                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"ziheng_193_-40_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)            
            
            
            
            ############################ Ziheng 193 198_-60 #################################
            ############################################################ Determine excel file path##############################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'test193-198.xlsx')

            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame

            # DataFrame -----> Python List
            data_value = df.values.tolist()
            #####################################################################################################################################
            #generate variable list
            list_P = (data_value[2])[1:]
            list_Q = (data_value[3])[1:]
            list_Rs = (data_value[0])[1:]
            list_Xs = (data_value[1])[1:]


            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas

            
            # Use canvas controller to find components by name
            Q = canvas.find("master:const", "Q_Setting") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")

            
            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")

            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Disabled"
            figure9.state = "Disabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Disabled"
            table13minus40.state = "Disabled"
            table1360.state = "Disabled"
            table13minus60.state = "Enabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Disabled"

            # Run each case
            for index in range(len(list_Rs)):
                
                # Change variables each cycle
                Q.parameters(Name="Q_Setting", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])


                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"ziheng_193_-60_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)

            ################################# Ziheng 155-160 #####################################
            #################################################  Determine excel file path #############################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'test155-160.xlsx')

            #Read data in excel 
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame

            # DataFrame -----> Python List
            data_value = df.values.tolist()
            ##########################################################################################################################
            #generate variable list
            list_P = (data_value[2])[1:]
            list_Q = (data_value[3])[1:]
            list_Rs = (data_value[0])[1:]
            list_Xs = (data_value[1])[1:]


            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas
            
            # Use canvas controller to find components by name
            Q = canvas.find("master:const", "Q_Setting") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")

            
            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")

            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Enabled"
            figure8.state = "Disabled"
            figure9.state = "Disabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Disabled"
            table13minus40.state = "Disabled"
            table1360.state = "Disabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Disabled"

            # Run each case (168)
            for index in range(len(list_Rs)):
                
                # Change variables each cycle
                Q.parameters(Name="Q_Setting", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])


                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"ziheng_155_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)

            ################################################### Ziheng 178-181 ########################################
            ############################################################### Determine excel file path #############################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'test178-181.xlsx')

            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame

            # DataFrame -----> Python List
            data_value = df.values.tolist()
            #######################################################################################################################################
            
            #generate variable list
            list_P = (data_value[2])[1:]
            list_Q = (data_value[3])[1:]
            list_Rs = (data_value[0])[1:]
            list_Xs = (data_value[1])[1:]


            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas
            canvas1 = FaultTest.canvas("Grid_Side_Ctrl") # get the controller of grid side controller canvas
            
            # Use canvas controller to find components by name
            Q = canvas.find("master:const", "Q_Setting") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")
            
            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")

            # Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Enabled"
            figure9.state = "Disabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Disabled"
            table13minus40.state = "Disabled"
            table1360.state = "Disabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Disabled"
            
            # Run each case (168)
            for index in range(len(list_Rs)):
                
                # Change variables each cycle
                Q.parameters(Name="Q_Setting", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])


                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"ziheng_178_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)

            ######################################## Ziheng 182-185 ###########################################
            ######################################################## Determine excel file path ###################################################
            path = r'C:\Users\Niu2021\Desktop\integration\input_data'
            xlsPath = os.path.join(path,'test182-185.xlsx')

            ## Read data in excel
            df = pd.read_excel(xlsPath) # default read the firt sheet in Excel and save as a DataFrame

            # DataFrame -----> Python List
            data_value = df.values.tolist()
            #######################################################################################################################################
            #generate variable list
            list_P = (data_value[2])[1:]
            list_Q = (data_value[3])[1:]
            list_Rs = (data_value[0])[1:]
            list_Xs = (data_value[1])[1:]


            # Select the specific component
            canvas = FaultTest.canvas("Main") # get the controller of main canvas
            canvas1 = FaultTest.canvas("Grid_Side_Ctrl") # get the controller of grid side controller canvas
            
            # Use canvas controller to find components by name
            Q = canvas.find("master:const", "Q_Setting") #in grid-side controller
            P = canvas.find("master:const", "P_Setting")
            Rs = canvas.find("master:const", "Rgrid")
            Xs = canvas.find("master:const", "Xgrid")

            
            # Select the layer (enabled/disabled)
            TOVLay = FaultTest.layer("TOV_layer")
            figure23 = FaultTest.layer("figure23")
            figure8 = FaultTest.layer("figure8")
            figure9 = FaultTest.layer("figure9")
            figure10green = FaultTest.layer("figure10green")
            figure10blue = FaultTest.layer("figure10blue")
            figure10red = FaultTest.layer("figure10red")
            figure111hz = FaultTest.layer("figure111hz")
            figure1110hz = FaultTest.layer("figure1110hz")
            figure6blue = FaultTest.layer("figure6blue")
            figure6orange = FaultTest.layer("figure6orange")
            figure6grey = FaultTest.layer("figure6grey")
            figure6yellow = FaultTest.layer("figure6yellow")
            figure7blue = FaultTest.layer("figure7blue")
            figure7yellow = FaultTest.layer("figure7yellow")
            table1340 = FaultTest.layer("table1340")
            table13minus40 = FaultTest.layer("tableminus1340")
            table1360 = FaultTest.layer("table1360")
            table13minus60 = FaultTest.layer("tableminus1360")
            Yuxiangtest = FaultTest.layer("Yuxiangtest")
            Muyuantest = FaultTest.layer("Muyuantest")
            NonMuyuan = FaultTest.layer("NonMuyuan")
            large1to120 = FaultTest.layer("large1to120")

            #Layer Settings
            TOVLay.state = "Disabled"
            figure23.state = "Disabled"
            figure8.state = "Disabled"
            figure9.state = "Enabled"
            figure10green.state = "Disabled"
            figure10blue.state = "Disabled"
            figure10red.state = "Disabled" 
            figure111hz.state = "Disabled" 
            figure1110hz.state = "Disabled" 
            figure6blue.state = "Disabled" 
            figure6orange.state = "Disabled" 
            figure6grey.state = "Disabled" 
            figure6yellow.state = "Disabled" 
            figure7blue.state = "Disabled" 
            figure7yellow.state = "Disabled"
            table1340.state = "Disabled"
            table13minus40.state = "Disabled"
            table1360.state = "Disabled"
            table13minus60.state = "Disabled"
            Yuxiangtest.state = "Disabled"
            Muyuantest.state = "Disabled"
            NonMuyuan.state = "Enabled"
            large1to120.state = "Disabled"

            # Run each case (168)
            for index in range(len(list_Rs)):
                
                # Change variables each cycle
                Q.parameters(Name="Q_Setting", Value=list_Q[index])
                P.parameters(Name="P_Setting", Value=list_P[index])
                Rs.parameters(Name="Rgrid", Value=list_Rs[index])
                Xs.parameters(Name="Xgrid", Value=list_Xs[index])


                # Saving the output file
                FaultTest.parameters(PlotType="1", output_filename=f"ziheng_182_{index+1}.out")
                FaultTest.run()


                messages = FaultTest.messages()
                for msg in messages:
                    print("%s  %s  %s" % (msg.scope, msg.status, msg.text))

                print("-"*60)
                output = FaultTest.output()
                print(output)

    finally:
        # Exit PSCAD
        pscad.quit()

else:
    LOG.error("Failed to launch PSCAD")