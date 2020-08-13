import rhinoscriptsyntax as rs
import Eto
import Rhino
import scriptcontext as sc

__commandname__ = "DefineExcelOutput"

class DisplayField:
    def __init__(self,label,sheet,range,form):
        self.label=label
        self.sheet=sheet
        self.range=range
        self.form=form
        self.action=''
        self.eto=None
    def getEtoRow(self):
        if self.eto:
            self.label=self.eto[0].Text
            self.sheet=self.eto[1].Text
            self.range=self.eto[2].Text
        self.eto=[
            Eto.Forms.TextBox(Text=self.label),
            Eto.Forms.TextBox(Text=self.sheet),
            Eto.Forms.TextBox(Text=self.range),
        
            Eto.Forms.Button(Text='Up'),
            Eto.Forms.Button(Text='Down'),
            Eto.Forms.Button(Text='Delete')
            ]
        
        self.eto[3].Click+=self.onUp
        self.eto[4].Click+=self.onDown
        self.eto[5].Click+=self.onDelete
        
        for i in range(3):
            self.eto[i].Size=Eto.Drawing.Size(75,30)
        for i in range(3,6):
            self.eto[i].Size=Eto.Drawing.Size(40,20)
    
        self.etoTableCells=[]
        for item in self.eto:
            self.etoTableCells.append(Eto.Forms.TableCell(item,scaleWidth=True))
        return Eto.Forms.TableRow(self.etoTableCells)
        
    def getList(self):
        output=[]
        for i in range(3):
            output.append(self.eto[i].Text)
        return output
    def onUp(self,sender,e):
        if self.form:
            self.action='up'
            self.form.redraw()
            self.action=''
    def onDown(self,sender,e):
        if self.form:
            self.action='down'
            self.form.redraw()
            self.action=''
    def onDelete(self,sender,e):
        if self.form:
            self.action='delete'
            self.form.redraw()
            self.action=''
        
class Dialog_WindowName(Eto.Forms.Dialog):
    def __init__(self):
        self.Title="Edit Output Fields"
        self.Padding=Eto.Drawing.Padding(15)
        self.Resizable=True
        
        self.Key=DisplayField("Label","Sheet","Range",False)
        if not "displayFields" in sc.sticky:
            sc.sticky["displayFields"]=[]
        self.DisplayFieldLists=sc.sticky["displayFields"]
        self.DisplayFieldEtos=[]
        for item in self.DisplayFieldLists:
            self.DisplayFieldEtos.append(DisplayField(item[0],item[1],item[2],self))
        
        self.Button_Add=Eto.Forms.Button(Text='Add')
        self.Button_Add.Click+=self.Add
        self.Button_OK=Eto.Forms.Button(Text='Save')
        self.Button_OK.Click+=self.OnOKButtonClick
        self.Button_Cancel=Eto.Forms.Button(Text='Cancel')
        self.Button_Cancel.Click+=self.OnCancelButtonClick
        
        self.draw()
        self.Update=False
        
    def draw(self):
        
        layout = Eto.Forms.TableLayout()
        layout.Spacing = Eto.Drawing.Size(10,10)
        layout.Rows.Add(self.Key.getEtoRow())
        for item in self.DisplayFieldEtos:
            layout.Rows.Add(item.getEtoRow())
        layout.Rows.Add( Eto.Forms.TableRow(
            Eto.Forms.TableCell(self.Button_Add,scaleWidth=False),
            Eto.Forms.TableCell(self.Button_Cancel, scaleWidth=False), 
            Eto.Forms.TableCell(self.Button_OK, scaleWidth=False)) )
        self.Content = layout
    
    def redraw(self):
        i=0
        while i<len(self.DisplayFieldEtos):
            if self.DisplayFieldEtos[i].action=='':
                i+=1
            elif self.DisplayFieldEtos[i].action=='up':
                if i>0:
                    self.DisplayFieldEtos.insert(i-1,self.DisplayFieldEtos.pop(i))
                i+=1
            elif self.DisplayFieldEtos[i].action=='down':
                if i+1<len(self.DisplayFieldEtos):
                    self.DisplayFieldEtos.insert(i+2,self.DisplayFieldEtos.pop(i))
                self.DisplayFieldEtos[i].action=''
            elif self.DisplayFieldEtos[i].action=='delete':
                self.DisplayFieldEtos.pop(i)
        self.draw()
    
    def Add(self,sender,e):
        self.DisplayFieldEtos.append(DisplayField("","","",self))
        self.redraw()
    
    def GetUserInput(self):
        self.DisplayFieldLists=[]
        for item in self.DisplayFieldEtos:
            self.DisplayFieldLists.append(item.getList())
        return self.DisplayFieldLists
    def OnCancelButtonClick(self, sender, e):
        print 'Canceled...'
        self.Close()
    
    def OnOKButtonClick(self, sender, e):
        print 'Applying Changes'
        self.Update = True
        self.Close()
    
    def GetUpdateStatus(self):
        return self.Update

def RunCommand( is_interactive ):
  print "Hello", __commandname__
  # get a point
  dialog = Dialog_WindowName()
  rc = dialog.ShowModal(Rhino.UI.RhinoEtoApp.MainWindow)
  
  if dialog.GetUpdateStatus():
      sc.sticky["displayFields"]=dialog.GetUserInput()
  
  return 0
RunCommand(True)