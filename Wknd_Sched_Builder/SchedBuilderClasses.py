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
        self.assignments=None
        self.skills=skills 


class Schedule():
    def __init__(self,slots,ee,preAssn,senList,polling):
        self.slots= slots  #A collection of Slot objects that compose this schedule
        self.oslots= deepcopy(slots)
        self.ee=ee
        self.preAssn=preAssn
        self.senList=senList
        self.polling=polling

    def evalAssnList(self):
        """Enter all predefined assignments into the schedule"""
        for myAssn in self.preAssn:
            if myAssn[0]==1: #Only evaluate those with '1' in 'Active' (first) column
                if myAssn[5]=="" or myAssn[5]==None: #If no job assigned, then this is likely a DNS determination. Assign to all jobs in slot.
                    myKeys=[]
                    for seqNo in range(myAssn[2],myAssn[3]+1):#Apply to all slots in given range.. +1 due to range fn not being inclusive
                        plus=[k for k in self.slots.keys() if k[:len(str(seqNo))+1]==str(seqNo)+'_'] #Pull dict keys for SLots where it is a slot with matching seqNo, regardless of job name
                        myKeys.extend(plus)
                    return myKeys