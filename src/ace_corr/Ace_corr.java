/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package ace_corr;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.Vector;
// import per excel

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;


/**
 *
 * @author pr41103
 */
public class Ace_corr {
   

    /**
     * @param args the command line arguments
     */
    public static FilenameFilter setfiltro (final String estensione){
         final String extension = "."+estensione;
         FilenameFilter textFilter = new FilenameFilter() {
			public boolean accept(File dir, String name) {
				String lowercaseName = name.toLowerCase();
				if (lowercaseName.endsWith(extension)) {
					return true;
				} else {
					return false;
				}
			}
		};
         return textFilter;
    }
    
    public static File[] listafile (String percorso,String estensione){
        File folder = new File(percorso);
        //Creazione filtro
                       
        File[] listOfFiles = folder.listFiles(setfiltro(estensione));
        
        //Scrive il nome file per controllare
        for (int j = 0; j < listOfFiles.length; j++) {
          if (listOfFiles[j].isFile()) {
            //System.out.println("File " + listOfFiles[j].getName());
          } else if (listOfFiles[j].isDirectory()) {
            //System.out.println("Directory " + listOfFiles[j].getName());

          }
        }
        
        return listOfFiles;  
    }
    
    public static File[] listafile (String percorso){
        return listafile (percorso,0);
    }
    public static File[] listafile (String percorso,int rest){
        File[] file = null;
        File[] directory = null;
        File folder = new File(percorso);
        File[] listOfFiles = folder.listFiles();
        int f=0,d=0;
        //Scrive il nome file per controllare
        if (rest !=0) {
            for (int j = 0; j < listOfFiles.length; j++) {
              if (listOfFiles[j].isFile()) {
                //System.out.println("File " + listOfFiles[j].getName());
                f++;
              } else if (listOfFiles[j].isDirectory()) {
                //System.out.println("Directory " + listOfFiles[j].getName());
                d++;
              }
            }
              file =new File[f];
              directory=new File[d];
              f=0;
              d=0;
              for (int j = 0; j < listOfFiles.length; j++) {
              if (listOfFiles[j].isFile()) {
                //System.out.println("File " + listOfFiles[j].getName());
                file[f]=listOfFiles[j];
                f++;
              } else if (listOfFiles[j].isDirectory()) {
                //System.out.println("Directory " + listOfFiles[j].getName());
                directory[d]=new File(percorso);
                directory[d]=listOfFiles[j];
                d++;
              }
            }
        }
        
        //System.out.println("fine");
        
        switch (rest) {
            case 0 : return listOfFiles;
            case 1 : return file;
            case 2 : return directory;
        }
        return listOfFiles;
    }
    public static String leggixls (String path,String nome) throws IOException{
        String[] Correzione=new String[0];
        return leggixls (path,nome,Correzione);
        
    }
    public static String leggixls (String path,String nome,String[]Correzione) throws IOException
    {
        //Vector<String> Prova=new Vector<String>[9];
        String pathun,pathor,dati;
        double tmpval=0;
        double tmpcor=0;
        double delta=0;
        String cella="";
        Boolean corr = Correzione.length>1;
        dati=nome+";";
        //prova[0]=
        pathor=path+"\\"+nome;//creo il percorso del file excel
        //pathun=path+"\\un"+nome;
        
        //DA RIPRISTINARE
        //System.out.println("elaborazione "+pathor);
        
        //System.out.println(pathun);
        //leggo il file excel
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(pathor));
        //leggo il foglio
        HSSFSheet s = workbook.getSheet("Dispersioni");
        
        //Pareti opache verso l'esterno
        cella = valorecella(s,"m",39);
        dati+=cella+";";
        if (corr){
            dati+=calcolodelta(cella,Correzione[1])+";";
        }else dati+=";";
        
        //Serramenti
        cella = valorecella(s,"m",40);
        dati+=cella+";";
        if (corr){
           dati+=calcolodelta(cella,Correzione[3])+";";     
        }else dati+=";";
        
        //Ponti termici
        cella=valorecella(s,"m",41);
        dati+=cella+";";
        if (corr){
           dati+=calcolodelta(cella,Correzione[5])+";";
        }else dati+=";";
        
        //Attraverso il terreno
        cella=valorecella(s,"m",42);
        dati+=cella+";";
        if (corr){
           dati+=calcolodelta(cella,Correzione[7])+";";
        }else dati+=";";
        
        //Verso ambienti non riscaldati
        cella=valorecella(s,"m",43);
        dati+=cella+";";
        if (corr){
           dati+=calcolodelta(cella,Correzione[9])+";";
        }else dati+=";";
        
        //Per ventilazione
        cella=valorecella(s,"m",44);
        dati+=cella+";";
        if (corr){
           dati+=calcolodelta(cella,Correzione[11])+";";
        }else dati+=";";
        
        //Totale disp.
        cella=valorecella(s,"m",46);
        dati+=cella+";";
        if (corr){
           dati+=calcolodelta(cella,Correzione[13])+";";
        }else dati+=";";
        
        //Classe
        s = workbook.getSheet("6 Risultati          CRTN");
        dati+=valorecella(s,"u",1);
        dati+=valorecella(s,"w",2).replaceAll("0\\.0","")+";";
        
        
        System.out.println(dati);
        return dati;
    }
    public static String calcolodelta (String valore,String corretto){
        String retdelta="";
        double tmpval=0;
        double tmpcor=0;
        double delta=0;
        boolean errore = false;
        try 
        {
            //((tema - corretto)/corretto)*100
            tmpval = Double.parseDouble(valore);
        }
        catch (NumberFormatException e)
        {
            errore=true;
            System.out.println("errore: Valore="+valore);
        }
        try 
        {
            //((tema - corretto)/corretto)*100
            tmpcor = Double.parseDouble(corretto);
        }
        catch (NumberFormatException e)
        {
            errore=true;
            System.out.println("errore: Corretto="+corretto);
        }
        if  (tmpval==tmpcor){
            retdelta="0";
            errore=true;
        }
        if (!errore && tmpcor !=0){
                delta = ((tmpval-tmpcor)/tmpcor)*100;
                retdelta+=delta;
        }
        return retdelta;
    }
    public static String valorecella (HSSFSheet s,String col,int row  ){
        String caratteri="abcdefghijklmnopqrstuvwxyz";
        int colonna=caratteri.indexOf(col);
        //System.out.println("colonna"+col+" "+colonna);
        String dati="";
        //imposto la riga -1
        HSSFRow srow = s.getRow(row-1);
        //imposto la colonna -1 es M = 12 a=0
        HSSFCell cell = srow.getCell(colonna);
         switch (cell.getCellType()) {
            case HSSFCell.CELL_TYPE_NUMERIC:
                double nval=cell.getNumericCellValue();
                //System.out.println(nval);
                dati+=nval;
            break;
            case HSSFCell.CELL_TYPE_STRING:
                String sval=cell.getStringCellValue();
                //System.out.println(sval);
                dati+=sval;
                //System.out.printlncell.getStringCellValue());
            case HSSFCell.CELL_TYPE_FORMULA:
                if (cell.getCachedFormulaResultType()==HSSFCell.CELL_TYPE_NUMERIC){
                    double fnval=cell.getNumericCellValue();
                    //System.out.println(fnval);
                    dati+=fnval;
                }
                if (cell.getCachedFormulaResultType()==HSSFCell.CELL_TYPE_STRING){
                    HSSFRichTextString fsval=cell.getRichStringCellValue();
                    //System.out.println("Last evaluated as \"" + fsval + "\"");
                    dati+=fsval;
                break; 
                }
                
            break;

            default: break;
        }
        return dati;
    }
    
    public static void scrivifile(String path,String nome,String riga){
        scrivifile (path,nome,riga,true);         
    }
     
    public static void scrivifile(String path,String nome,String riga,boolean append){
        
        try {
          String pathfile=path+"\\"+nome;
          File file = new File(pathfile);
          FileWriter Output = new FileWriter(file,append);
          PrintWriter outFile = new PrintWriter(Output);            
          outFile.println(riga);
          outFile.close();
        } catch (IOException e) {
          System.out.println("Errore: " + e);
          System.exit(1);
        }
    }
    
    /**
     *
     */
    public static Sessione[] elencoconsessione ()
    {
        String path="C:\\temiace\\temi";
        File[] lista = listafile(path,2);
        Sessione[] anno2013=new Sessione[lista.length];;
        int i=0;
        for (File testo:lista){
            anno2013[i]=new Sessione();
            String tmpnome=testo.getName();
            anno2013[i].nome=tmpnome;
            anno2013[i].corretto=listafile(testo.getPath(),"xls");
            anno2013[i].aule=creaaule(testo.getPath());
            i++;
            
        }
        System.out.println("fine sessione");
        return anno2013;
    }
    
    public static Aula[] creaaule(String path){
        File[] lista = listafile(path,2);
        Aula[] aule =new Aula[lista.length];
        int i=0;
        for (File testo:lista){
            aule[i]=new Aula();
            aule[i].nome=testo.getName();
            aule[i].lista=listafile(testo.getPath(),"xls");
            i++;
        }
        System.out.println("fine aule");
        return aule;
    }
    public static void main(String[] args) throws IOException {
            String path="C:\\temiace\\temi";
            Sessione[] Anno2013=elencoconsessione();
            String riga="";
            riga+="traccia;";
            riga+="aula;";
            riga+="nome;";
            riga+="Pareti opache verso l'esterno;";
            riga+="delta Pareti opache verso l'esterno;";
            riga+="Serramenti;";
            riga+="delta Serramenti;";
            riga+="Ponti termici;";
            riga+="delta Ponti termici;";
            riga+="Attraverso il terreno;";
            riga+="delta Attraverso il terreno;";
            riga+="Verso ambienti non riscaldati;";
            riga+="delta Verso ambienti non riscaldati;";
            riga+="Per ventilazione;";
            riga+="delta Per ventilazione;";
            riga+="Totale disp.;";
            riga+="delta Totale disp.;";
            riga+="Classe;";
            scrivifile(path,"risultatidef.csv",riga,false);
            String[] Divisa= riga.split("\\;");
            String[] Correzione=null;
            for (Sessione tmpsess:Anno2013){
                String rigas=tmpsess.nome+";";
                if (tmpsess.corretto.length==1){
                    String tmpstr = leggixls (tmpsess.corretto[0].getParent(),tmpsess.corretto[0].getName());
                    Correzione = tmpstr.split("\\;");
                    tmpstr=tmpstr.replaceAll("\\.xls","" );
                    tmpstr=tmpstr.replaceAll("\\.","," );
                    scrivifile(path,"risultatidef.csv",rigas+"corretto;"+tmpstr);
                }
                for (Aula tmpaula:tmpsess.aule){
                    String rigaa=tmpaula.nome+";";
                    for (File tmpfile:tmpaula.lista){
                        String rigaf=leggixls (tmpfile.getParent(),tmpfile.getName(),Correzione);
                        rigaf=rigaf.replaceAll("\\.xls","" );
                        rigaf=rigaf.replaceAll("\\.","," );
                        scrivifile(path,"risultatidef.csv",rigas+rigaa+rigaf);
                    }
                        
                }
            }
            /*
            String path="C:\\temiace\\temi";
            File[] lista = listafile(path);
            for (File testo:lista){
                File[] lista1=listafile(testo.getPath());
            }
            */
        
            /*
            
            for (File testo:lista){
                System.out.print(" path "+testo.getPath());
                System.out.print(" parent "+testo.getParent());
                System.out.println(" name "+testo.getName());
                riga=leggixls (testo.getParent(),testo.getName());
                riga=riga.replaceAll("\\.xls","" );
                riga=riga.replaceAll("\\.","," );
                
                System.out.println(riga);
                scrivifile(path,"risultatidef.csv",riga);
            }
            */
    }
}
