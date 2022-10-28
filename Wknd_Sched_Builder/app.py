import gradio as gr
from datetime import datetime
import SchedBuilderUtyModule as tls
import SchedBuilderClasses2 as cls
#Just in case:
import openpyxl as pyxl
import pandas as pd
import numpy as np
from copy import deepcopy

#######################
#First define the functions
def PrimeVisualTemplate(vTmplFl,FTrefFl,TempRefFl):
    primedTemplateName=tls.translate_Visual_Template(vTmplFl,FTrefFl,TempRefFl)
    return primedTemplateName,primedTemplateName,datetime.now(),datetime.now(),FTrefFl.name,TempRefFl.name #Return twice so as to send file to two interface entities (file holders), plus time stampers, return the refusal sheets to send to B

def GenerateSchedule(schedWWF,xtraDays,wkHrs,DaysCrew,AssnFl,FTrefFl,TempRefFl,PollFl,stop=None,template=False):
    assnWWF=schedWWF
    if assnWWF=='Yes': assnWWF=True
    if assnWWF=='No': assnWWF=True
    xtraDays=xtraDays
    Acrew=DaysCrew
    wkHrs=wkHrs
    mySched=tls.preProcessData(Acrew,wkHrs,FTrefFl,TempRefFl,AssnFl,PollFl,assnWWF=assnWWF,xtraDays=xtraDays)
    mySched.evalAssnList()
    mySched.proofEligVol()
    mySched.proofEligVol()#i don't know man i found in testing sometimes it wasnt completing properly the first time around
    if template==True: #Case of template viewer
        flNm=sch.printToExcel()
        return flNm,datetime.now()
    sch=mySched.fillOutSched_v3(stop=stop) #Stop is not None when submitted from tab C
    flNm=sch.printToExcel()
    return flNm,datetime.now()


#######################
#Second Defining the Interface
with gr.Blocks() as demo:
    with gr.Tab("A - Visual Template Builder"):
        gr.Markdown("On this tab you can input a purely visual template (assignment tables empty) and have it primed for use in the scheduling algorithm. The info generated here will also be sent as inputs to tab B. See the documentation for more details on required formatting")
        with gr.Row():
            A_fl_FTref=gr.File(label="Full Time Refusals Sheet")
            A_fl_Tref=gr.File(label="Temp Refusal Sheet")
        with gr.Row():
            A_fl_VT=gr.File(label="Visual Template File")
            with gr.Column():
                A_fl_PT=gr.File(label="Primed Template File")
                A_tx_PTtimestamp=gr.Textbox(label="File Change Timestamp - Primed Template",placeholder="Waiting for file")
        A_bt_PT = gr.Button("Prime Visual Template")
    with gr.Tab("B - Scheduler"):
        gr.Markdown("On this tab the inputs are used to generate a weekend schedule. The Template and refusal sheet files are carried over (with slightly garbled names) when generated using tab A. The schedule is considered 'primed' when the tables on the secondary sheets match what is indicated on the visual template. The other inputs on this sheet should be selected per the circumstances. For example. If Friday is part of the weekend to be scheduled, Select '32 hrs', 'Friday', and 'yes to WWF scheduling'. Normal non-long weekend will not assign OT to WWF, and be 40 hour weeks without Friday or Monday to be scheduled.")
        with gr.Row():
            B_fl_FTref=gr.File(label="Full Time Refusals Sheet")
            B_fl_Tref=gr.File(label="Temp Refusal Sheet")
        with gr.Row():
            with gr.Column():
                B_fl_PT=gr.File(label="Primed Template File")
                B_tx_PTtimestamp=gr.Textbox(label="File Change Timestamp - Primed Template",placeholder="Waiting for file")               
            with gr.Column():
                with gr.Row():
                     B_fl_Pl=gr.File(label="Polling File")
                with gr.Tab("Non-File Inputs"):
                    with gr.Row():
                        B_wwfOT=gr.Radio(["Yes", "No"],label="Assign OT to WWF? (If 'yes', WWF will be considered for filling in slots beyond their prescribed shifts in the Assignment List)")
                        B_dayCrew=gr.Radio(["Bud","Blue"],label="Which crew is on A shift this week?")
                    with gr.Row():
                        B_wkHrs=gr.Radio([32, 40],label="Regular Work Hours This Week?")
                        B_xtraDay=gr.CheckboxGroup(["Friday", "Monday"], label="Check the boxes as appropriate if scheduling long weekend")
        B_bt_MS = gr.Button("Generate Schedule")
        B_fl_FS=gr.File(label="Generated Schedule")
        B_tx_FTtimestamp=gr.Textbox(label="File Change Timestamp - Completed Schedule",placeholder="Waiting for first run")   
    with gr.Tab("C - Review"):
        with gr.Tab("Inspect Template"):
            gr.Markdown("On this tab you can have the program generate a schedule using only the template file so as to confirm the template was entered correctly for program interpretation. This process builds the template using all inputs present in tab B. Inputs are required for building the template.")
            C_bt_MT = gr.Button("Generate Template")
            C_fl_T=gr.File(label="Generated Template")
            C_tx_Ttimestamp=gr.Textbox(label="File Change Timestamp - Generated Template",placeholder="Waiting for first run.")  
        with gr.Tab("Force Stop Mid-Scheduling"):
            gr.Markdown("On this tab you can observe what a schedule looked like after making a specific number of assignments. Iteration number can be retrieved from the bottom of the verbose assignment list tab of a generated schedule.")
            with gr.Row():
                with gr.Column():
                    C_tx_FS = gr.Textbox(label="Time Of Full Schedule Generation",placeholder="N/A")
                    C_nm_PS = gr.Number(label="Limit Number for Total Assignments")
                with gr.Column():
                    C_fl_PS=gr.File(label="Partially Complete Schedule")
            C_bt_PS = gr.Button("Partially Generate Schedule")
        
    #######################
    #Third Define the Interactions
    A_bt_PT.click(PrimeVisualTemplate,[A_fl_VT,A_fl_FTref,A_fl_Tref],[A_fl_PT,B_fl_PT,A_tx_PTtimestamp,B_tx_PTtimestamp,B_fl_FTref,B_fl_Tref])
    B_bt_MS.click(GenerateSchedule,[B_wwfOT,B_xtraDay,B_wkHrs,B_dayCrew,B_fl_PT,B_fl_FTref,B_fl_Tref,B_fl_Pl],[B_fl_FS,B_tx_FTtimestamp])
    C_bt_MT.click(GenerateSchedule,[B_wwfOT,B_xtraDay,B_wkHrs,B_dayCrew,B_fl_PT,B_fl_FTref,B_fl_Tref,B_fl_Pl],[C_fl_T,C_tx_Ttimestamp])
    demo.launch()
