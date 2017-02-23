/**
 * Created by dmitry on 15.10.16.
 */
import java.io.*;
import java.util.ArrayList;
import jxl.*;
import jxl.Workbook;
import jxl.write.*;

import static java.lang.System.in;

public class Main {
    static public void loadNames(String path , ArrayList <String> l) throws IOException {
        File file = new File(path);
        try {
            InputStream in = new FileInputStream(file);
            InputStreamReader read = new InputStreamReader(in);
            BufferedReader br = new BufferedReader(read);
            String n = new String();
            while(n != null) {
                n = br.readLine();
                if(n != null)
                    l.add(n);
            }
        } catch (IOException e) {
            System.out.println("Exception");
        }
    }
    public static void loadWorkbook(String path , ArrayList <String> names , ArrayList <Workbook> wb) throws Exception
    {
        int i = 1;
        try {
            for(i = 0; i < names.size(); i++) {
                System.out.println(path + names.get(i));
                String s = names.get(i).trim();
                Workbook tmp = Workbook.getWorkbook(new File(path + s));
                wb.add(tmp);
            }
        }
        catch(Exception ex)
        {
            System.out.println(ex.getMessage());
        }
    }
    public static  void writeIntoCell(int x , int y, String data, WritableWorkbook w_wb) throws Exception
    {
        WritableSheet w_sh = w_wb.getSheet(0);
        try
        {
            Label l = new Label(x,y,data);
            w_sh.addCell(l);
        }
        catch(Exception ex)
        {
            ex.printStackTrace();
        }
    }
    public static void rewriteRow(int amount, int row ,int counter, int w_row, Workbook wb, WritableWorkbook w_wb) throws Exception
    {
        WritableSheet w_sh = w_wb.getSheet(0);
        Sheet sh = wb.getSheet(0);
        try{
            for(int i = 0; i < amount ; i ++){
                Cell c = sh.getCell(i,row);
                Label l;
                if(i == 0) {
                     l = new Label(i, w_row, String.valueOf(counter));
                }
                else
                {
                     l = new Label(i,w_row,c.getContents());
                }
                w_sh.addCell(l);
            }
        }
        catch (Exception ex){
            ex.printStackTrace();
        }

    }
    public static int  read_req(ArrayList<String> req) throws Exception
    {
        int a = 0;
        try
        {
            BufferedReader br_mon = new BufferedReader(new InputStreamReader(new FileInputStream(new File("/home/dmitry/Projects/Source_xls/req_month"))));
            BufferedReader br_years = new BufferedReader(new InputStreamReader(new FileInputStream(new File("/home/dmitry/Projects/Source_xls/req_years"))));
            ArrayList<String> years = new ArrayList<>();
            ArrayList<String> mon = new ArrayList<>();
            String s = new String();
            String s2 = new String();
            while (s!=null || s2!= null)
            {
                s = br_mon.readLine();
                s2 = br_years.readLine();
                if(s != null)
                {
                    s.trim();
                    mon.add(s);
                }
                if(s2 != null)
                {
                    s2.trim();
                    years.add(s2);
                }
            }
            a = years.size();
            for(int i = 0; i < mon.size() ;i ++)
            {
                for(int j = 0; j < years.size() ; j++)
                {
                    req.add("." + mon.get(i) +"." + years.get(j));
                }
            }
        }
        catch(Exception ex)
        {
            ex.printStackTrace();
        }
        return a;
    }
    public static void print_row(Workbook wb, int amount , int y)
    {
        Sheet sh = wb.getSheet(0);
        for(int i = 0; i < amount ; i++)
        {
            Cell c = sh.getCell(i,y);
            System.out.print(c.getContents() + "  ");
        }
    }
    public static void processingXls(ArrayList<WritableWorkbook> w_wb_a, ArrayList <Workbook> wb, ArrayList<String> name)
    {
        try {
            WritableWorkbook w_wb = w_wb_a.get(0);
            int counter = 0;
            int w_row = 0;
            ArrayList req = new ArrayList<String>();
            int years_size = 0;
            years_size = read_req(req);
            System.out.println(years_size);
            for(int q = 0; q < req.size() ; q ++) {
                if(q % years_size == 0) {
                    //if(q != 0)
                    //    w_row += 10;
                    w_row = 0;
                    counter = 0;
                    int w = q / years_size;
                    w_wb = w_wb_a.get(w);
                    String month = new String();
                    switch(w){
                        case 0: month = "Январь";
                            break;
                        case 1 : month = "Февраль";
                            break;
                        case 2 : month = "Март";
                            break;
                        case 3 : month = "Апрель";
                            break;
                        case 4 : month = "Май";
                            break;
                        case 5 : month = "Июнь";
                            break;
                        case 6 : month = "Июль";
                            break;
                        case 7 : month = "Август";
                            break;
                        case 8 : month = "Сентябрь";
                            break;
                        case 9 : month = "Октябрь";
                            break;
                        case 10 : month = "Ноябрь";
                            break;
                        case 11 : month = "Декабрь";
                            break;
                    }
                    WritableSheet w_sh_tmp = w_wb.getSheet(0);
                    Label tmp = new Label(1, w_row,month);
                    w_sh_tmp.addCell(tmp);
                    w_row ++;
                }
            for (int i = 0; i < wb.size(); i++) {
                Workbook tmp = wb.get(i);
                Sheet sh = tmp.getSheet(0);
                int j = 0;
                System.out.println("\n" + name.get(i));

                    for (j = 1; j < sh.getRows(); j++) {
                        Cell c = sh.getCell(4, j);
                        String s = c.getContents();
                        //System.out.println(s);
                        String p = (String) req.get(q);
                        //System.out.println(j);
                       // print_row(tmp,9,j);
                        if (s.contains(p)){
                            counter ++;
                            rewriteRow(9,j,counter,w_row,tmp,w_wb);
                            w_row ++;
                            //System.out.println(" ");
                        }
                    }
                }
            }
        }
        catch(Exception ex)
        {
            ex.printStackTrace();
        }
    }
    public static void main(String args[]) {
        ArrayList  names = new ArrayList<String>();
        ArrayList wb = new ArrayList <Workbook> ();
        ArrayList <WritableWorkbook> w_wb_a = new ArrayList<>();
        //ArrayList<String> output_names = new ArrayList<>();
        try {
            loadNames("/home/dmitry/Projects/Source_xls/names", names);
            for(int i = 0; i < names.size() ; i++)
                System.out.println(names.get(i));
            for(int i = 0; i < 12; i++)
            {
                String p = "/home/dmitry/Projects/Parsed_XLS/pased_" + i + ".xls";
                OutputStream os = new FileOutputStream(new File(p));
                WritableWorkbook w_wb = Workbook.createWorkbook(os);
                w_wb.createSheet("1",0);
                w_wb_a.add(w_wb);
            }

            loadWorkbook("/home/dmitry/Projects/Source_xls/",names,wb);
            processingXls(w_wb_a,wb,names);
            /*int max_len = 0;
            for(int i = 0; i < w_wb.getSheet(0).getColumns(); i++) {
                max_len = 0;
                for(int j = 0; j < w_wb.getSheet(0).getRows(); j++)
                {
                    int len = w_wb.getSheet(0).getCell(i,j).getContents().length();
                    if(len > max_len)
                        max_len = len;
                }
                CellView cv = w_wb.getSheet(0).getColumnView(i);
                cv.setSize(max_len*340);
                w_wb.getSheet(0).setColumnView(i, cv);
            }
            */
            for(int q = 0; q < 12; q++){
                WritableWorkbook w_wb = w_wb_a.get(q);
                int max_len = 0;
                for(int i = 0; i < w_wb.getSheet(0).getColumns(); i++) {
                    max_len = 0;
                    for(int j = 0; j < w_wb.getSheet(0).getRows(); j++)
                    {
                        int len = w_wb.getSheet(0).getCell(i,j).getContents().length();
                        if(len > max_len)
                            max_len = len;
                    }
                    CellView cv = w_wb.getSheet(0).getColumnView(i);
                    cv.setSize(max_len*340);
                    w_wb.getSheet(0).setColumnView(i, cv);
                }
                for(int o = 0 ; o < w_wb.getSheet(0).getRows(); o++){
                    CellView c = w_wb.getSheet(0).getRowView(o);
                    c.setSize(600);
                    w_wb.getSheet(0).setRowView(o,c);
                }
                w_wb.write();
                System.out.println(w_wb.getNumberOfSheets());
                w_wb.close();
            }


        }
        catch (Exception ex)
        {
            System.out.println(ex.getMessage());
        }
    }
}
