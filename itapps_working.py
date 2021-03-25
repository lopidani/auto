#!/usr/bin/env python2
import pymysql,datetime,yagmail,os,sys,requests
from reportlab.platypus import SimpleDocTemplate,Table,TableStyle,Paragraph,Spacer,PageBreakIfNotEmpty,Image
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfgen import canvas

#------ RELEASE TYPE --------
#release="production"
release="testing no email"
#release="testing yes email"
#----------------------------
# perioada verificarilor in zile
period=5
# date calendaristice
now= datetime.date.today().strftime('%d.%m.%Y')
# from now + period
fromnowp=(datetime.date.today()+datetime.timedelta(days=period)).strftime('%d.%m.%Y')
# paragraf cu data
p="Data verificarii: {}".format(now) 
# titluri tabele expirate
t1='Auto care au ITP-ul expirat'
t2='Auto care au RCA-ul expirat'
t3='Auto care au CASCO-ul expirat'
t4='Auto care au ROVIGNIETA expirata'
t5='Auto care au LEASING-ul expirat'
# titluri tabele neexpirate
t6='Auto la care ITP-ul va expira peste {} zile'.format(period)
t7='Auto la care RCA-ul va expira peste {} zile'.format(period)
t8='Auto la care CASCO-ul va expira peste {} zile'.format(period)
t9='Auto la care ROVIGNIETA va expira peste {} zile'.format(period)
t10='Auto la care LEASING-ul va expira peste {} zile'.format(period)
# capete de coloana tabele
c0=" Nr.";c1 = "Nr auto";c2 = "Marca auto";c3 = "Sofer auto";c4 = " Data Itp   "
c5 = " Data Rca   ";c6 = " Data Casco  ";c7 = " Data Rovignieta";c8 = "Data Leasing"
# EXPIRATE (OUT)
head_itp =[c0,c1,c2,c3,c4]
out_itp=list()
out_itp.append(head_itp)
head_rca=[c0,c1,c2,c3,c5]
out_rca=list()
out_rca.append(head_rca)
head_casco=[c0,c1,c2,c3,c6]
out_casco=list()
out_casco.append(head_casco)
head_rov=[c0,c1,c2,c3,c7]
out_rov=list()
out_rov.append(head_rov)
head_leas=[c0,c1,c2,c3,c8]
out_leas=list()
out_leas.append(head_leas)
# NEEXPIRATE (IN)
in_itp=list()
in_itp.append(head_itp)
in_rca=list()
in_rca.append(head_rca)
in_casco=list()
in_casco.append(head_casco)
in_rov=list()
in_rov.append(head_rov)
in_leas=list()
in_leas.append(head_leas)   

# pdf-ul nostru ete de forma: auto_05.03.2021.pdf
file= "auto_{}.pdf".format(now)
pdf = SimpleDocTemplate(file, pagesize=A4)
styles=getSampleStyleSheet()
styleN = styles["Normal"]
head_style1 = styles['Heading5']
head_style2 = styles['Heading6']
elem = []

# database settings
host='192.168.20.35'
user='root'
passw='put password'
database='auto'
chs='utf8mb4'
db_tb='sheet1'

class AUTO(object):
      def __init__(self,host,user,passw,chs,database):
          self.host=host
          self.user=user
          self.passw=passw
          self.chs=chs
          self.database=database
          # datas
          self.out_itp=out_itp
          self.out_rca=out_rca
          self.out_casco=out_casco
          self.out_rov=out_rov
          self.out_leas=out_leas
          self.in_itp=in_itp
          self.in_rca=in_rca
          self.in_casco=in_casco
          self.in_rov=in_rov
          self.in_leas=in_leas
          # variable date expirate sau nu
          self.exp=True
          self.neexp=True
          self.network=True

      def working_directory(self,wd,icon,log):
          if sys.platform=="win32":
             sep="\\" 
          elif sys.platform=="linux2":
               sep="/"   
          os.chdir(wd)
          ico=sep.join((wd,icon))
          # log file
          self.lf=sep.join((wd,log))
          # logo
          self.image = Image(icon,13.5*cm,2*cm)
          self.image.hAlign = 'LEFT'
          # inainte de fiecare rulare emplty log.txt
          self.clear_logFile(self.lf)   
                           
      def interogate(self):
          try:
              con = pymysql.connect(host=host,user=user,password=passw,charset=chs,database=database)
              with con.cursor() as cur:
                   cur.execute("SELECT  NrCirculatie,MarcaTip,Sofer,DataITP,DataRCA,DataCasco,DataRovignieta,DataLeasing from %s"%(db_tb))  
                   dta = cur.fetchall()
                   i=j=k=l=m=n=o=r=s=t=0
                   for d in dta :
                       try:                      
                            # date expirate la itp
                            if (d[3]-datetime.date.today()).days< 0:
                                i+=1
                                self.out_itp.append([i,d[0],d[1],d[2],d[3]])  
                            elif (d[3]-datetime.date.today()).days ==period: 
                                j+=1   
                                self.in_itp.append([j,d[0],d[1],d[2],d[3]])
                            # date expirate la rca         
                            if (d[4]-datetime.date.today()).days< 0:
                                k+=1
                                self. out_rca.append([k,d[0],d[1],d[2],d[4]])
                            elif (d[4]-datetime.date.today()).days== period: 
                                l+=1   
                                self.in_rca.append([l,d[0],d[1],d[2],d[4]])
                            # date expirate la casco             
                            if (d[5]-datetime.date.today()).days< 0:
                                m+=1
                                self.out_casco.append([m,d[0],d[1],d[2],d[5]])  
                            elif (d[5]-datetime.date.today()).days== period:
                                n+=1    
                                self.in_casco.append([n,d[0],d[1],d[2],d[5]]) 
                            # date expirate la rovignieta                        
                            if (d[6]-datetime.date.today()).days< 0:
                                o+=1
                                self.out_rov.append([o,d[0],d[1],d[2],d[6]])
                            elif (d[6]-datetime.date.today()).days==period:
                                r+=1     
                                self.in_rov.append([r,d[0],d[1],d[2],d[6]])
                            # date expirate la leasing                
                            if (d[7]-datetime.date.today()).days< 0:
                                s+=1
                                self.out_leas.append([s,d[0],d[1],d[2],d[7]])
                            elif (d[7]-datetime.date.today()).days==period:
                                t+=1    
                                self.in_leas.append([t,d[0],d[1],d[2],d[7]])                            
                       except TypeError: pass             
          except Exception as e:con=False;self.edit_errors("Function interogate ==> {}".format(str(e))+'\n',self.lf)
          if con:con.close() 
          # EXPIRATE
          if (len(self.out_itp)<=1) and (len(self.out_rca)<=1) and (len(self.out_casco)<=1) and (len(self.out_rov)<=1) and (len(self.out_leas)<=1):
              self.exp=False
          # NEEXPIRATE
          if (len(self.in_itp)<=1) and (len(self.in_rca)<=1) and (len(self.in_casco)<=1) and (len(self.in_rov)<=1) and (len(self.in_leas)<=1):
              self.neexp=False      
              
      def draw_pdf(self):
          try:
            # Table style
            self.TbS = TableStyle([
                                    ('TEXTCOLOR',(0,0),(-1,-1),colors.black),
                                    ('ALIGN',(0,0),(-1,-1),'CENTER'),
                                    ('FONTNAME',(0,0),(4,0),'Courier-Bold'),
                                    ('FONTSIZE',(0,0),(4,0),10),
                                    ('BOTTOMPADDING',(0,0),(-1,0),10),
                                    # linii orizontale
                                    ('LINEBELOW',(0,0),(-1,-1),1,colors.black),
                                    ('GRID',(0,0),(-1,-1),1,colors.black), 
                                    # linia exterioara (contur)
                                    ('BOX',(0,0),(4,-1),1,colors.black),
                                    ('BACKGROUND',(0,1),(-1,-1),colors.white),
                                    ('TEXTCOLOR',(0,0),(-1,0),colors.white)
                                 ])
            
            # EXPIRATE
            if self.exp==False: 
               elem.append(Paragraph('Nu avem auto cu date expirate !', head_style1))
               elem.append(Spacer(0.2*cm, 0.3* cm))
            else:
                if len(self.out_itp) >1:
                   self.build(elem,out_itp,head_style1 ,head_style2,t1,colors.brown)
                if len(self.out_rca) >1:
                   self.build(elem,out_rca,head_style1 ,head_style2,t2,colors.brown)
                if len(self.out_casco) >1:
                   self.build(elem,out_casco,head_style1 ,head_style2,t3,colors.brown)
                if len(self.out_rov) >1:
                   self.build(elem,out_rov,head_style1 ,head_style2,t4,colors.brown)                    
                if len(self.out_leas) >1:
                   self.build(elem,out_leas,head_style1 ,head_style2,t4,colors.brown)

            # NEEXPIRATE 
            if self.neexp==False: 
                elem.append(Paragraph( 'Pe {} nu avem auto cu date expirate !'.format(fromnowp),head_style1))
                elem.append(Spacer(0.2*cm, 0.3* cm))
            else:
                if len(self.in_itp) >1:
                   self.build(elem,in_itp,head_style1 ,head_style2,t6,colors.green)                        
                if len(self.in_rca) >1:
                   self.build(elem,in_rca,head_style1 ,head_style2,t7,colors.green)                                    
                if len(self.in_casco) >1:
                   self.build(elem,in_casco,head_style1 ,head_style2,t8,colors.green)  
                if len(self.in_rov) >1:
                   self.build(elem,in_rov,head_style1 ,head_style2,t9,colors.green) 
                if len(self.in_leas) >1:
                   self.build(elem,in_leas,head_style1 ,head_style2,t10,colors.green)       
                          
            def draw_pag_nr(canvas,doc):
                canvas.saveState()
                canvas.setFont('Times-Roman',5)
                canvas.drawString(2.7*cm, 1* cm,"pag. %d" % doc.page)
                canvas.setFont('Times-Bold',6.5)
                canvas.drawString(17*cm,29* cm,"Departamentul IT&C")
                canvas.drawString(17*cm,28.7* cm,"GSP Offshore Dana 34")
                canvas.restoreState()    
            pdf.build(elem,onFirstPage=draw_pag_nr,onLaterPages=draw_pag_nr)
          except Exception as e:self.edit_errors("Function draw_pdf ==> {}".format(str(e))+'\n',self.lf)    

      def build(self,list1,list2,style1,style2,title,color):
          list1.append(PageBreakIfNotEmpty())
          list1.append(self.image)
          list1.append(Spacer(0.2*cm, 0.3* cm))       
          list1.append(Paragraph(title,style1))
          list1.append(Spacer(0.2*cm, 0.3* cm))
          list1.append(Paragraph(p,style2))
          tb = Table(list2, repeatRows=1)
          tb.hAlign = 'LEFT'
          tb_style=self.TbS
          tb_style.add('BACKGROUND',(0,0),(7,0),color)
          tb.setStyle(tb_style)
          list1.append(tb)       

      def send_email(self):
          try:  
             smtp = yagmail.SMTP(user="gsp.itapps@gmail.com",
                                 password="K2kb2bdbgv!k",
                                 host='smtp.gmail.com')                    
             # email subject
             subject = 'Situatia auto la data: '+now
             # email content with attached file from current dir.
             content = ['Trimitem in format pdf situatia auto pentru GSP ,',
                        'asa cum reiese din baza de date in {}.'.format(now),
                        'Verificarea se face in fiecare zi iar emailul se trimite,',
                        'numai daca peste {} zile sunt date care expira.'.format(period),   
                        'Persoana de contact : daniel.lopatica@gspoffshore.com',
                        file]            
             # send email
             if release=="production":
                principal='irina.apostol@gspoffshore.com'
                cc=['george.badescu@gspoffshore.com','catalin.ionescu@gspoffshore.com','daniel.lopatica@gspoffshore.com']
                if self.neexp==True : 
                   smtp.send(to=principal,cc=cc,subject=subject,contents=content) 
             elif release == "testing yes email":
                  principal='daniel.lopatica@gspoffshore.com'
                  cc=[]
                  if self.neexp==True :
                     smtp.send(to=principal,cc=cc,subject=subject,contents=content)
          except Exception as e:
                 self.edit_errors("Function send_mail ==> {}".format(str(e))+'\n',self.lf)

      # write errors on hdd (create log.txt only if errors occurs)
      def edit_errors(self,logStr,logFile):    
          with open(logFile,'a+') as lfa: 
               lfa.write(logStr)

      def clear_logFile(self,logFile):    
          if os.path.exists(logFile)==True:
             with open(logFile,'w') as lfw:
                  pass             

      def check_network(self):
          try:
              request = requests.get("https://dns.google",timeout=2)
          except (requests.ConnectionError,requests.Timeout):
                 self.network=False
                 self.edit_errors("Function check_network ==> {}".format("Nu avem internet !")+'\n',self.lf)
                                                 
      def clear_pdf(self):
          # since pdf file is personalized we must delete the file
          if release != "testing no email":    
             try:
                 os.remove(file)
             except OSError: pass
       
if __name__=="__main__":
   # Instantiate AUTO class    
   IA=AUTO(host,user,passw,chs,database)
   # windows location 
   if sys.platform=="win32":
      wd="C:\\Users\\daniel.lopatica\\Desktop\\GitHub\\auto"  
   # linux server location
   elif sys.platform=="linux2":
        wd="linux location"
   icon="gsp.jpg"
   log="log.txt"              
   IA.working_directory(wd,icon,log)
   IA.check_network()
   if IA.network:
      IA.interogate()
      IA.draw_pdf()
      IA.send_email()
      IA.clear_pdf()  
