
class Slot():
    """A single 4 hour time slot for a single job, to be filled by 1 person"""
    def __init__(self,seqID,dispNm,trnNm=None):
        self.trnNm=trnNm #to be used when filtering out staff for training
        self.dispNm=dispNm #to be used for printouts
        self.seqID=seqID
        #self.datetime=   #Determined based on seq. Used in printout of assignments
        self.assignee=None #To store eeid of assignee for printout\
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
    def __init__(self,slots):
        #self.Slots=   #A collection of Slot objects that compose this schedule
        pass

    def preFill():
        """Enter all predefined assignments into the schedule"""