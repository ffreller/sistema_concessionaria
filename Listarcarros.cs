using System;
using System.IO;
using NetOffice.ExcelApi;

namespace sistema_concessionaria{
    
    public class ListarCarros
    {
        public void Listarcarros()
        {
        Application ex = new Application();
        ex.Workbooks.Open(@"C:\Users\40809588897\Desktop\Programar\Semana 5\sistema_concessionaria\vendas.xls");
        int contador = 1;
        do
        {
            Console.WriteLine("Dados do Cliente:");
            Console.WriteLine(ex.Cells[contador,1].Value + "; " + ex.Cells[contador,2].Value + "; " + ex.Cells[contador,3].Value + "; " + ex.Cells[contador,4].Value + "; " + ex.Cells[contador,5].Value + "; " + ex.Cells[contador,6].Value);
            Console.WriteLine("Dados do Carro:");
            Console.WriteLine(ex.Cells[contador,7].Value + "; " + ex.Cells[contador,8].Value + "; " + ex.Cells[contador,9].Value + "; Ar-condicionado: " + ex.Cells[contador,10].Value + "; Airbag: " + ex.Cells[contador,11].Value + "; ABS:" + ex.Cells[contador,12].Value);
            Console.WriteLine(ex.Cells[contador,13].Value + " reais" + " em " + ex.Cells[contador,14].Value + " parcelas");
            contador += 1;
        }
        while(ex.Cells[contador,1].Value != null);
        
        ex.ActiveWorkbook.Save();
        ex.Quit();
        ex.Dispose();
        }
    }
}