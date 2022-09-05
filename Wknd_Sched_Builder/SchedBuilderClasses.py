from copy import deepcopy

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
    
    def key(self):
        return str(self.seqID)+'_'+self.dispNm
    
    def assn(self,sch,assnType=None,slAssignee=None):
        """Assign a slot to someone"""
        self.assnType=assnType 
        self.assignee=slAssignee #eeid
        if slAssignee is not None:
            sch.ee[slAssignee].assignments[self.key()]=self #add this slot to the ee's assigned slot dictionary
        del sch.oslots[self.key()] #Remove this slot from the 'openslots' collection

class ee():
    """A staff persons data as related to weekend scheduling"""
    def __init__(self,snty,crew,id,Last,First,refHrs,wkHrs,wkndHrs=0,skills=[]):
        self.seniority=snty
        if refHrs==None:
            self.refHrs=0
        else:
            self.refHrs=refHrs
        self.wkdyHrs=wkHrs
        self.wkndHrs=wkndHrs
        self.lastNm=Last
        self.firstNm=First
        self.eeID=id
        self.crew=crew
        self.assignments={}#To be appended with slots as they are assigned, keyed as they are in the slot dictionary
        self.skills=skills 


class Schedule():
    def __init__(self,slots,ee,preAssn,senList,polling):
        self.slots= slots  #A collection of Slot objects that compose this schedule
        self.oslots= deepcopy(slots)
        self.ee=ee #A dictionary containing ee info
        self.preAssn=preAssn #A list of lists containing the predefind assignment info
        self.senList=senList
        self.polling=polling

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
                if assnTp=='DNS':#Case of 'Do Not Schedule' modification to a slot
                    if (myAssn[5]=="" or myAssn[5]==None): jb=None #grab job name
                    else: jb=myAssn[5]
                    keys=getKeys(myAssn[2],myAssn[3],jb) #pull all the keys for slots this particular assn list item applies to
                    for k in keys:
                        slChLg.append([k,self,'DNS',None]) #Add record(s) to the slot change long, one for each record.
                elif assnTp=='WWF': #Case of WWF assignment being specified in the assignment log
                    jb=myAssn[5]
                    keys=getKeys(myAssn[2],myAssn[3],jb) #pull all the keys for slots this particular assn list item applies to
                    for k in keys:
                        slChLg.append([k,self,'WWF',myAssn[4]]) #include eeid in the assigment info...
                elif
