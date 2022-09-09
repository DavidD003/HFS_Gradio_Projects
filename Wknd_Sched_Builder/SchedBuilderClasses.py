from copy import deepcopy
from ctypes import alignment
from operator import index
from typing import KeysView
import openpyxl as pyxl
import functools

def debug(func):
    """Print the function signature and return value"""
    @functools.wraps(func)
    def wrapper_debug(*args, **kwargs):
        args_repr = [repr(a) for a in args]                      # 1
        kwargs_repr = [f"{k}={v!r}" for k, v in kwargs.items()]  # 2
        signature = ", ".join(args_repr + kwargs_repr)           # 3
        print(f"Calling {func.__name__}({signature})")
        value = func(*args, **kwargs)
        print(f"{func.__name__!r} returned {value!r}")           # 4
        return value
    return wrapper_debug


class Slot():
    """A single 4 hour time slot for a single job, to be filled by 1 person"""
    def __init__(self,seqID,dispNm,trnNm=None):
        self.trnNm=trnNm #to be used when filtering out staff for training
        self.dispNm=dispNm #to be used for printouts
        self.seqID=seqID
        #self.datetime=   #Determined based on seq. Used in printout of assignments
        self.assignee=None #To store eeid of assignee for printout
        self.assnType=None #e.g. Forced/WWF
        self.slotInShift=0 #1 if first slot in someones shift. 2 if second, etc.
        self.totSlotsInShift=0 # 1 if 4 hours shift, 2 if 8 hour shift, 3 if 12 hour shift
        self.eligVol=0 #This will be used to track which slot is the most constrained...it tracks the count of how many people eligible volunteers this slot has going for it
        self.disallowed=[] #EEID's that were specified as not allowed to be assigned to this slot in the Assn List

    def key(self):
        return str(self.seqID)+'_'+self.dispNm
    
    
    def assn(self,sch,assnType=None,slAssignee=None,fromList=False):
        """Assign a slot to someone, and perform associated variable tracking etc."""
        if (slAssignee is not None) and assnType=='DNS': #Case that this is specifying *not* to assign someone. In every other case it is a matter of actually assigning someone 
            self.disallowed.append(slAssignee)
        else:
            self.assnType=assnType 
            self.assignee=slAssignee #eeid
            del sch.oslots[self.key()] #Remove this slot from the 'openslots' collection
            if slAssignee is not None: #Case of specific assignment, only not follwoed through when its no ee and DNS
                sch.ee[slAssignee].assnBookKeeping(self,sch) #add this slot to the ee's assigned slot dictionary & other tasks
        #Logging for printout after
        logTxt=''
        if fromList==True:
            logTxt+= 'Per Assn List: '
        if assnType=='DNS':
            logTxt+='Removed slot '+self.dispNm+' '+ sch.slLeg[self.seqID-1][2]+' ('+sch.slLeg[self.seqID-1][1]+') from scheduling'
            if slAssignee is not None: logTxt+= ' for ee '+ str(slAssignee)
        elif assnType=='WWF': logTxt+="WWF Assignment: EE "+ str(slAssignee)+' to ' +self.dispNm+' '+ sch.slLeg[self.seqID-1][2]+' ('+sch.slLeg[self.seqID-1][1]+')'
        elif assnType=='F': logTxt+="FORCED Assignment: EE "+ str(slAssignee)+' to ' +self.dispNm+' '+ sch.slLeg[self.seqID-1][2]+' ('+sch.slLeg[self.seqID-1][1]+')'
        elif assnType=='V': logTxt+="Voluntary Assignment: EE "+ str(slAssignee)+' to ' +self.dispNm+' '+ sch.slLeg[self.seqID-1][2]+' ('+sch.slLeg[self.seqID-1][1]+')'
        elif assnType=='N':logTxt+="No voluntary or forced assignment could be made to "+self.dispNm+' '+ sch.slLeg[self.seqID-1][2]+' ('+sch.slLeg[self.seqID-1][1]+')'
        sch.assnLog.append(logTxt)

class ee():
    """A staff persons data as related to weekend scheduling"""
    def __init__(self,snty,crew,id,Last,First,refHrs,wkHrs,wkndHrs=0,skills=[]):
        self.seniority=int(snty)
        if refHrs==None:
            self.refHrs=0
        else:
            self.refHrs=float(refHrs)
        self.wkdyHrs=float(wkHrs)
        self.wkndHrs=float(wkndHrs)
        self.lastNm=Last
        self.firstNm=First
        self.eeID=int(id)
        self.crew=crew
        self.assignments=[]#To be appended with slots as they are assigned, keyed as they are in the slot dictionary
        self.skills=skills 
    
    def dispNm(self,slt=None):
        if slt is None:
            if int(self.seniority)>50000: return self.firstNm[0]+'.'+self.lastNm[0]+self.lastNm[1:].lower()+'(T)' #Case of Temps
            else: return self.firstNm[0]+'.'+self.lastNm[0]+self.lastNm[1:].lower()
        else: #Can make functionality to pass in 'MESR' to display name based on some slot criteria?
            pass

    
    def assnBookKeeping(self,sl,sch):
        """Carried out when assigned to slot, adjusts tally of eligible volunteers to other slots accordingly"""
        self.assignments.append(sl.key())
        self.wkndHrs+=4
        #   Return keys for open slots where slot is same time as one just assigned and its for a job the person is trained on
        kys=[k for k in sch.oslots.keys() if (k[len(str(sl.seqID))-1]==str(sl.seqID) and k[len(str(sl.seqID)):] in self.skills)]
        for k in kys:
            sch.oslots[k].eligVol.pop(sch.oslots[k].eligVol.index(self.eeID)) #pop the eeId out of eligVol list for relevant slots

    
    def totShiftHrs(self,sl):
        """Given a slot, assuming it is assigned, what is the total shift length of the shift in which that slot is a constituent"""
        sLen=1 #start off shift length at one because the slot being passed in is always minimum
        assnSeqIDs=[int(k[:k.index('_')]) for k in self.assignments] #Pull out the seqID's form the key strings for each slot an ee is already assigned
        if len(assnSeqIDs)==0:
            return 4 #Case that no slots assigned so far, so if sl assigned then its alone
        else:
            anch=sl.seqID
            assnSeqIDs.append(anch)
            assnSeqIDs.sort() #sorts integers lowest to highest
            for offset in [-3,-2,-1,1,2,3]:
                if anch+offset in assnSeqIDs: #Knowing that shifts will never be assigned beyond 3 in a row, need only check 4 in a row in the event this function is being run testing for the 4th being ok. Need to test for 4 in a row to rule it out
                    sLen+=1
            return sLen*4 #Return num hours
    
    
    def assnConflict(self,sl):
        """Returns true if someone is already assigned to a slot with same seqID as potential assignment, false if no conflict"""
        assns=[int(k[:k.index('_')]) for k in self.assignments] #Pull out the seqID's form the key strings for each slot an ee is already assigned
        if len(assns)==0:return False #if no other assignments.. no conflict
        elif sl.seqID not in assns: return False #if no other assns with same seqID.. no conflict
        elif sl.seqID in assns: return True #if other assignment already amde with same seqID... true, conflict present

    
    def gapOK(self,sl,sch,tp='V'):
        """Returns true if the slot, when assigned, doesn't break the rule for minimum gap between shifts. 12 hours for forcing, 8 for vol"""
        assns=[int(k[:k.index('_')]) for k in self.assignments] #Pull out the seqID's form the key strings for each slot an ee is already assigned
        #Just need to check the nearest neighbours aren't with a gap of 1 empty slot
        deltaNextShifts=[v for v in [sId-sl.seqID for sId in assns] if v>0]  # Define this and the next before defining next/prevNeighDist to manage case of 0 this being equal to [], which min() can't accept
        deltaPrevShifts=[v for v in [sl.seqID-sId  for sId in assns] if v>0] ##Basically, subtract various Id's from sl in question. Remove vlaues<0 bc those slots come later (next neighb). Then take the minimum value, which is the diff in seqID between sl being compared, and nearest neighbour on left when sorted chronologically 
        if len(deltaNextShifts)==0:
            deltaNextShifts=[0] #Fill with bogus # to give 'ok' result to function
        if len(deltaPrevShifts)==0:
            deltaPrevShifts=[0]
        nextNeighbDist=min(deltaNextShifts)
        prevNeighbDist=min(deltaPrevShifts) 
        #Utility functions:
        def okForLastWkShift():
            if self.crew==sch.Bcrew and sl.seqID-6*(sch.friOT==False)<3+1*(tp=='F'):
                return False #Case of B shift worker being assigned within 8 (or 12 if forced) hrs of last weeks final afternoon shift, no go
            elif self.crew==sch.Acrew and tp=='F' and sl.seqID-6*(sch.friOT==False)<2:
                return False #Case of A shift worker being forced onto first slot of the weekend
            else: return True
        def okForNextWkShift():
            if self.crew=='Rock' and sl.seqID+6*(sch.monOT==False)==23:
                return False #Case of night shift person being assigned 3p-7pafternoon before weekday night shift
            elif self.crew==sch.Acrew and tp=='F' and sl.seqID+6*(sch.monOT==False)==24:
                return False #Case of day shift ee (next week) being forced 7p-11p the night before
            else: return True
        #=====
        if tp=='V' and (nextNeighbDist==2 or prevNeighbDist==2):
            return False#Distance of one means consecutive slots (shiftlength already checked), distance greater than 2 is gap of 8 hours or more
        elif tp=='F' and ((nextNeighbDist==2 or prevNeighbDist==2) or prevNeighbDist==3):return False #Forcing requires gap of 12h or more going into it, 8 hours after it
        elif okForLastWkShift()==True and okForNextWkShift()==True: return True #Reaching this elif means that the other conditions aren't true, so lastly just have to check the gap with weekday shifts ok
        else: return False #Some condition not met       

    
    def slOK(self,sch,sl,poll=0,tp='V'):
        """Returns True if the slot being tested is ok to be assigned, false if not"""
        #Test all conditions (trained, wk hrs, consec shift, time between shifts, before making a branch to test willingness or not based on assignment type forced/voluntary)
        if (sl.dispNm in self.skills) and (self.eeID not in sl.disallowed): #the person is trained and hasn't been specified in assignment log *not* to be assigned here
            if self.wkndHrs+self.wkdyHrs<=60: #total week hours ok!
                if self.assnConflict(sl)==False: #No existing assignment at same time!
                    if self.totShiftHrs(sl)<=12: #This slot wouldn't have a given shift exceed 12 hours
                        if self.gapOK(sl,sch,tp=tp): #This slot being assigned doesn't break a shift gap rule
                            if tp=='V':#voluntary: check willigness
                                if (poll[3+sl.seqID] !="") and (poll[3+sl.seqID] is not None) and (poll[3+sl.seqID]!='n'): #Person is willing!
                                    return True
                                else: return False
                            else: return True#Forced, no check on willingess
class Schedule():
    def __init__(self,Acrew,slots,ee,preAssn,senList,polling,slLeg):
        self.Acrew=Acrew
        if Acrew=='Bud':self.Bcrew='Blue'
        else:self.Bcrew='Bud'
        self.slots= slots  #A collection of Slot objects that compose this schedule
        self.oslots= deepcopy(slots)
        self.ee=ee #A dictionary containing ee info
        self.preAssn=preAssn #A list of lists containing the predefind assignment info
        self.senList=senList
        self.polling=polling
        self.assnLog=[] #To be appended when assignments made, for read out with final product
        self.slLeg=slLeg #Slot Legend. Used for easy refernece of slot times after.
        seqIDs=[int(k[:k.index('_')]) for k in self.slots]
        #look at slots to see if fri/mon OT to be assigned... used in 'gapOK' method for ee
        if min(seqIDs)>6: self.friOT=False
        else: self.friOT=True
        if max(seqIDs)<19: self.monOT=False
        else: self.monOT=True 
        self.rev=0

    
    def evalAssnList(self):
        """Enter all predefined assignments into the schedule"""
        #First, iterate through the assignment list. Each record will generate one or more records for the 'slot change log'
        #The slot change log will have one record for each slot for which an assignment (or other specified status change) should be made
        #So assignment log records that indicate a span of multiple slots will generate multiple records in the change log.
        slChLg=[] #Initialize slot change log. It will be a list of lists where each sublist has the necessary info within to pass to the slot assignment function
        #===
        def getKeys(seqId_1,seqId_2,jobNm=None):
            """Returns the keys for all slots that a given assnLog record applies to, job specified or not"""
            myKeys=[]
            if jobNm==None:
                for seqNo in range(seqId_1,seqId_2+1):#Apply to all slots in given range.. +1 due to range fn not being inclusive
                    moreKeys=[k for k in self.slots.keys() if k[:len(str(seqNo))+1]==str(seqNo)+'_'] #Pull dict keys for Slots where it is a slot with matching seqNo, regardless of job name
                    myKeys.extend(moreKeys)
            else: #job is defined
                for seqNo in range(seqId_1,seqId_2+1):
                    myKeys.append(str(seqNo)+'_'+jobNm)
            return myKeys
        #===
        #Here we get down to business - Reading the assignment log, generating records in the slotChangeLog, and then evaluating those changes
        for myAssn in self.preAssn:
            if myAssn[0]==1: #Only evaluate those with '1' in 'Active' (first) column
                assnTp=myAssn[1]
                if (myAssn[5]=="" or myAssn[5]==None): jb=None #grab job name
                else: jb=myAssn[5]
                if (myAssn[4]=="" or myAssn[4]==None): asgne=None #grab assignee
                else: asgne=myAssn[4]
                keys=getKeys(myAssn[2],myAssn[3],jb) #pull all the keys for slots this particular assn list item applies to
                for k in keys:
                    slChLg.append([k,self,assnTp,asgne]) #Add record(s) to the slot change long, one for each record.
        #Now that the slChLg is made, carry out the function that reads it record by record and goes and modifies the slots
        def evalLogRec(rec):
            """Carry out the 'assn' method on the associated slot with relevant data from Assn log"""
            self.slots[rec[0]].assn(rec[1],rec[2],rec[3],fromList=True)
        for rec in slChLg:
            evalLogRec(rec)
    
    
    def fillOutSched(self):
        """Having made the predetermined assignments, fill in the voids in the schedule"""
        #Algorithm is basically:
        # 1. Force staff for slots with no eligible assignees
        # 2. Iterate through unassigned slots in sequence of which is most constrained
        #    Assign staff in order of who gets priority pick at the slot
        # 3. Force for slots that had no eligible after giivng the eligble their voluntary choices
        #    If no forcing availability, label as such and move on
        #End when no more unassigned slots left
        
        def nextSlots(force=0):
            """Returns the next most constrained unassigned slot object. If 'forcing'=True then returns list of slots with 0 eligible assignes, ordered by seqID"""
            if force==0: #Proceed with selecting most constrained slot with >=1 potential assignees
                eligCnts=[len(s.eligVol) for s in self.oslots if len(s.eligVol)>0]
                if eligCnts.count(min(eligCnts))>1: #Case that there are slots tied for most constrained
                    slts=[s for s in self.oslots if len(s.eligVol)==min(eligCnts)] #Retrieve the tied slots
                    totSkills=[sum([len(self.ee[eId].skills) for eId in s.eligVol ]) for s in slts ] #For each slot, make a list of integers, whee each integer is the number of jobs an ee eligible for that slot is trained on. Sum those lists, and the slot with the highest number is selected, since that is correlated with the eligible assignees for that slot having the most ability to cover other slots.
                    if totSkills.count(max(totSkills))>1: #Case that 2 slots are tied for total # of training records for eligible operators
                        #Go with the one for which there is an operator with least spots trained... assuming that the operator who is most constrained training wise gets it.. while this is an assumption without great basis, there is at least the point that someone with less training will likely have less refusal hours in the year.. so it may turn to work out ok
                        trainRecForLeastTrainedEE=[min([len(self.ee[eId].skills) for eId in s.eligVol ]) for s in slts ] #Same formula as totSkills except min instead of sum
                        return slts[trainRecForLeastTrainedEE.index(min(trainRecForLeastTrainedEE))] #Here only 1 return statement because if its a tie we'll just take the first one, which the index function here will give.
                    else: return self.oslots[totSkills.index(max(totSkills))] #Case of one slot having more totSkills than another
                else:
                    return self.oslots[eligCnts.index(max(eligCnts))] #Retrieve most constrained if not tied with any other
            elif force==1:#Forcing for the first time. Return list of slots to force into in chronological order
                return sorted([self.oslots[s] for s in self.oslots if len(self.oslots[s].eligVol)==0],key=lambda x: x.seqID)
            elif force==2: #Forcing for teh 2nd time. Return all slots. the 'eligibility tracking' isn't perfect so can't filter by it
                return self.oslots

        def pickAssignee(sl,tp='V'):
            """Returns an eeid and the assignment type, either voluntary or forced, or 'N" for None/No staff, for the passed slot"""
            if tp=='V':
                def tblSeq(sl,Acrew,Bcrew):
                    """Returns keys for retrieving poll data tables in sequence of priority assignment. With respect to shift, assignment priority goes in CABCA sequence, first FT then Temps"""
                    if sl.seqID in [1,2,7,8,13,14,19,20]: homeShift='C'
                    elif sl.seqID in [3,4,9,10,15,16,21,22]: homeShift='A'
                    else: homeShift='B'
                    keys=[]
                    seqStr='CABCA'
                    cD={'C':'Rock','A':Acrew,'B':Bcrew} #Crew Dict
                    for eeTp in ['FT','Temp']: #Go through all FT's before temps
                        for i in range(3):#code below uses variable homeShift in conjunction with stepping through seqStr to pull out crews in order of priority selection of OT
                            keys.append('tbl_'+cD[seqStr[seqStr.index(homeShift)+i]]+eeTp)
                    return keys
                #===
                ks=tblSeq(sl,self.Acrew,self.Bcrew)
                #Relying on the fact that the tables in the excel sheet were already sequenced in order of refusal hours...
                for k in ks:#Iterate through the tables in provided sequence
                    for rec in self.polling[k]: #Iterate through rows in table
                        if self.ee[rec[0]].slOK(self,sl,poll=rec):return rec[0],'V' #Person has been found 
                return None,'N'                                   
            elif tp=='F':
                for i in range(len(self.senList)-1,-1,-1): #Work way down seniority list
                    lowManID=int(self.senList[i][2])
                    if self.ee[lowManID].slOK(self,sl,tp='F'): return lowManID,'F'
                return None,'N'
        
        #Proceed with carrying out the algorithm:
        # 1. Initial Forcing
        sls=nextSlots(force=1)
        for s in sls:
            eId,tp=pickAssignee(s,tp='F')
        s.assn(self,assnType=tp,slAssignee=eId)
        # 2. Voluntary Filling
        sls=deepcopy(self.oslots)
        for i in range(len(sls)): #Iterate across all slots
            s=nextSlots(sls) #Pick most constraind in set
            sls.pop(s.key()) #Remove that one from set for next iteration
            eId,tp=pickAssignee(s)
            self.oslots[s.key()].assn(self,assnType=tp,slAssignee=eId) #Note the assign method acts on originall, not deepcopy, retrieved via key
        # 3. Final Forced Filling
        sls=nextSlots(force=2)
        for s in sls:
            eId,tp=pickAssignee(s,tp='F')
        s.assn(self,assnType=tp,slAssignee=eId)
      




    def printToExcel(self):
        """Print all slot assignments to an excel file for human-readable schedule interpretation"""
        #Define Cell styling function
        def styleCell(cl,clType,horizMergeLength=0):
                for i in range(horizMergeLength+1):
                    cl=cl.offset(0,i)
                    if clType=='hours':
                        cl.font=pyxl.styles.Font(bold=True,size=16,color="00FFFFFF")
                        cl.fill=pyxl.styles.PatternFill(fill_type="solid",start_color='00FF0000',end_color='00FF0000')
                        cl.alignment=pyxl.styles.Alignment(horizontal='center')
                        cl.border=pyxl.styles.Border(left=pyxl.styles.Side(border_style='thick'),
                        right=pyxl.styles.Side(border_style='thick'),
                        top=pyxl.styles.Side(border_style='thick'),
                        bottom=pyxl.styles.Side(border_style='thick'))
                    elif clType=='shift':
                        cl.font=pyxl.styles.Font(bold=True,size=16)
                        cl.alignment=pyxl.styles.Alignment(horizontal='center')
                        cl.border=pyxl.styles.Border(left=pyxl.styles.Side(border_style='thick'),
                        right=pyxl.styles.Side(border_style='thick'),
                        top=pyxl.styles.Side(border_style='thick'),
                        bottom=pyxl.styles.Side(border_style='thick'))
                    elif clType=='day':
                        cl.font=pyxl.styles.Font(bold=True,size=28)
                        cl.alignment=pyxl.styles.Alignment(horizontal='center')
                        cl.border=pyxl.styles.Border(left=pyxl.styles.Side(border_style='thick'),
                        right=pyxl.styles.Side(border_style='thick'),
                        top=pyxl.styles.Side(border_style='thick'),
                        bottom=pyxl.styles.Side(border_style='thick'))
                    elif clType=='DNS':
                        cl.fill=pyxl.styles.PatternFill(fill_type="solid",start_color='B2B2B2',end_color='B2B2B2')
                        cl.font=pyxl.styles.Font(bold=True,size=14)
                        cl.alignment=pyxl.styles.Alignment(horizontal='center')
                        cl.border=pyxl.styles.Border(left=pyxl.styles.Side(border_style='thin'),
                        right=pyxl.styles.Side(border_style='thin'),
                        top=pyxl.styles.Side(border_style='thin'),
                        bottom=pyxl.styles.Side(border_style='thin'))
                    elif clType=='WWF':
                        cl.fill=pyxl.styles.PatternFill(fill_type="solid",start_color='00B0F0',end_color='00B0F0')
                        cl.font=pyxl.styles.Font(bold=True,size=14,color="00FFFFFF")
                        cl.alignment=pyxl.styles.Alignment(horizontal='center')
                        cl.border=pyxl.styles.Border(left=pyxl.styles.Side(border_style='thin'),
                        right=pyxl.styles.Side(border_style='thin'),
                        top=pyxl.styles.Side(border_style='thin'),
                        bottom=pyxl.styles.Side(border_style='thin'))
                    elif clType=='F':
                        cl.fill=pyxl.styles.PatternFill(fill_type="solid",start_color='CC99FF',end_color='CC99FF')
                        cl.font=pyxl.styles.Font(bold=True,size=14)
                        cl.alignment=pyxl.styles.Alignment(horizontal='center')
                        cl.border=pyxl.styles.Border(left=pyxl.styles.Side(border_style='thin'),
                        right=pyxl.styles.Side(border_style='thin'),
                        top=pyxl.styles.Side(border_style='thin'),
                        bottom=pyxl.styles.Side(border_style='thin'))
                    elif clType=='N':
                        cl.font=pyxl.styles.Font(bold=True,size=16,color="00FFFFFF")
                        cl.fill=pyxl.styles.PatternFill(fill_type="solid",start_color='00FF0000',end_color='00FF0000')
                        cl.alignment=pyxl.styles.Alignment(horizontal='center')
                        cl.border=pyxl.styles.Border(left=pyxl.styles.Side(border_style='thick'),
                        right=pyxl.styles.Side(border_style='thick'),
                        top=pyxl.styles.Side(border_style='thick'),
                        bottom=pyxl.styles.Side(border_style='thick'))
                    elif clType=='jbNm':
                        cl.font=pyxl.styles.Font(bold=True,size=14)
                        cl.alignment=pyxl.styles.Alignment(horizontal='left')
                        cl.border=pyxl.styles.Border(left=pyxl.styles.Side(border_style='thin'),
                        right=pyxl.styles.Side(border_style='thin'),
                        top=pyxl.styles.Side(border_style='thin'),
                        bottom=pyxl.styles.Side(border_style='thin'))


        #Initial Setup
        self.rev+=1
        wb=pyxl.Workbook()
        dest_filename = 'Wknd Sched_Rev '+str(self.rev)+'.xlsx'
        ws = wb.active
        ws.title = "Full Schedule" #Title worksheet for printout
        ws['A1']='Rev '+str(self.rev)
        styleCell(ws['A1'],'shift')
        ws['A2']='Manager'
        styleCell(ws['A2'],'shift')
        ws.column_dimensions['A'].width =22.78
        dys=['Friday','Saturday','Sunday','Monday']
        shifts=['C','A','B']*4
        tSlots=['11p - 3a','3a - 7a','7a - 11a', '11a - 3p','3p-7p', '7p-11p']*4
        for i in range(0,24,6): #Print Days of week
            cl=ws.cell(column=2+i,row=1)
            styleCell(cl,'day',6) #Style all the cells within the merge
            ws.merge_cells(start_row=1, start_column=2+i, end_row=1, end_column=2+i+5)
            cl.value=dys.pop(0)
        for i in range(0,24,2): #Print the shift title
            cl=ws.cell(column=2+i,row=3)
            styleCell(cl,'shift',1)
            ws.merge_cells(start_row=3, start_column=2+i, end_row=3, end_column=2+i+1)
            cl.value=shifts.pop(0)
        for i in range(24): #Print the shift times row
            cl=ws.cell(column=2+i,row=4)
            cl.value=tSlots.pop(0)
            styleCell(cl,'hours')

        def styleNfill(cl,s):
            if s.assnType=='DNS': cl.value="N/A"
            elif s.assignee is not None: cl.value=self.ee[s.assignee].dispNm()
            else: 
                s.assnType='N'
                cl.value='NO STAFF'
            styleCell(cl,s.assnType)

        #==================
        #Enter shifts
        #The sequence of keys returned from self.slots.keys() is the same as the slots required were defined in "AllSlots"
        #They should all be in order which would allow for a naive method here but I will generalize it in case someone changes the template
        #And they are not in order after all.
        #The method here is to use a dictionary keyed by job name to track which row a job is printed out to,
        #and to add a job to the dictionary if/when it is encountered for the first time.
        jrD={} #job/Row Dict {job:row}
        r=5 #Start printing job slot info's at row 5 in excel
        for k in self.slots:
            s=self.slots[k] #Retrieve slot
            if s.dispNm not in jrD: #Add to dict if not in it
                jrD[s.dispNm]=r
                jbCl=ws.cell(row=r,column=1)
                jbCl.value=s.dispNm
                styleCell(jbCl,'jbNm')
                r+=1 #increment for next one to be observed
            cl=ws.cell(row=jrD[s.dispNm],column=1+s.seqID)
            styleNfill(cl,s)
            ws.column_dimensions[chr(65+s.seqID)].width = max(10.33,len(cl.value),ws.column_dimensions[chr(64+s.seqID)].width-5) #Widen column if new value is wider than any previously existing
        #=========================================
        #Lastly, merge contiguous shifts cells
        def numInARow(cl,n=1):
            """Given a cell, return the number of cells in a row have the same name in them."""
            nextval=cl.offset(0,1).value
            if nextval is None or nextval=='':
                return n
            elif nextval==cl.value:
                n=numInARow(cl.offset(0,1),n+1) #Recursive fn. If next cell matches current, use this function on the next one again
            return n
        #====== Proceed with above function to be used
        nSkip=0 #Initialize for skipping cells in this loop as applicable
        for rw in range(5,r):
            for i in range(2,26):
                if nSkip<1:
                    inArow=numInARow(ws.cell(row=rw,column=i))
                    nSkip+=(inArow-1)
                    ws.merge_cells(start_row=rw, start_column=i, end_row=rw, end_column=i+inArow-1)
                elif nSkip>0: 
                    nSkip=nSkip-1 #This facilitates *not* checking cells that have already been merged.. thats because if i,i+1,i+2 merged at i, then when loop increments to i+1, it would give an error when checking on i+2.. need to skip to i+3 after having merged i,+1,+2
        wb.save(filename = dest_filename)