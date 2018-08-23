using System;
using System.IO;
using NetOffice.ExcelApi;

namespace sistema_concessionaria{
    
    public class CadastrarCarro{
    
        public void Cadastrarcarro()
        
            {Console.WriteLine("Cadastro de carro");
            Console.WriteLine("Qual o modelo do carro?");
            string modelocarro = Console.ReadLine();
            Console.WriteLine("Qual o ano do carro?");
            string anocarro = Console.ReadLine();
            Console.WriteLine("Qual o preço do carro(sem opcionais)?");
            string precocarro = Console.ReadLine();
            Console.WriteLine("Opcionais");
            Opcionais opcionais1 = new Opcionais();
            do{ 
                Console.WriteLine("Ar-condicionado? (s ou n)");
                opcionais1.arcon = Console.ReadLine();
                }
            while(opcionais1.arcon != "s" && opcionais1.arcon != "n");
            do{ 
                Console.WriteLine("Airbag? (s ou n)");
                opcionais1.airbag = Console.ReadLine();
                }
            while(opcionais1.airbag != "s" && opcionais1.airbag != "n");
            do{
                Console.WriteLine("Freios ABS? (s ou n)");
                opcionais1.abs = Console.ReadLine();
                }
            while(opcionais1.abs != "s" && opcionais1.abs != "n");
            
            
        if(!File.Exists(@"C:\Users\40809588897\Desktop\Programar\Semana 5\sistema_concessionaria\carros.xls"))
        {
            Criarexcel(modelocarro, anocarro, precocarro, opcionais1);
        }
        else
        {
            Application ex = new Application();
            ex.DisplayAlerts = false;
            ex.Workbooks.Open(@"C:\Users\40809588897\Desktop\Programar\Semana 5\sistema_concessionaria\carros.xls");
            int contador = 1;
            do
            {
                contador += 1;

            } while (ex.Cells[contador,1].Value != null);
            ex.Cells[contador,1].Value = modelocarro;
            ex.Cells[contador,2].Value = anocarro;
            ex.Cells[contador,3].Value = precocarro;
            ex.Cells[contador,4].Value = opcionais1.arcon;
            ex.Cells[contador,5].Value = opcionais1.airbag;
            ex.Cells[contador,6].Value = opcionais1.abs;
            ex.ActiveWorkbook.Save();
            ex.Quit();
            ex.Dispose();
        }
    }
    public void Criarexcel(string modelocarro, string anocarro, string precocarro,  Opcionais opcionais1)
    {    
        Application ex = new Application();
        ex.Workbooks.Add();
        ex.Cells[1,1].Value = modelocarro;
        ex.Cells[1,2].Value = anocarro;
        ex.Cells[1,3].Value = precocarro;
        ex.Cells[1,4].Value = opcionais1.arcon;
        ex.Cells[1,5].Value = opcionais1.airbag;
        ex.Cells[1,6].Value = opcionais1.abs;

        ex.ActiveWorkbook.SaveAs(@"C:\Users\40809588897\Desktop\Programar\Semana 5\sistema_concessionaria\carros.xls");
        ex.Quit();
        ex.Dispose();
    }
            


            }   
    }
