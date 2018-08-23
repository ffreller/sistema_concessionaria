using System;
using System.IO;
using NetOffice.ExcelApi;
namespace sistema_concessionaria{
  
public class CadastrarCliente
{
    public void Cadastrarcliente()
    {
        Console.WriteLine("Cadastro");
        Console.WriteLine("Qual é seu nome?");
        string nome = Console.ReadLine();
        Console.WriteLine("Qual é seu e-mail?");
        string email = Console.ReadLine();
        Console.WriteLine("Qual é seu CPF/CNPJ?");
        string cpfecnpj = Console.ReadLine();
        Console.WriteLine("Qual sua cidade?");
        Endereco endereco1 = new Endereco();
        endereco1.cidade = Console.ReadLine();
        Console.WriteLine("Qual seu bairro?");
        endereco1.bairro = Console.ReadLine();
        Console.WriteLine("Qual sua rua?");
        endereco1.rua = Console.ReadLine();
        Console.WriteLine("Número:");
        endereco1.numero = Console.ReadLine();
        string cidade = Convert.ToString(endereco1.cidade);
        string bairro = Convert.ToString(endereco1.bairro);
        string rua = Convert.ToString(endereco1.rua);
        string numero = Convert.ToString(endereco1.numero);

       
        if(!File.Exists(@"C:\Users\40809588897\Desktop\Programar\Semana 5\sistema_concessionaria\clientes.xls"))
        {
            Criarexcel(nome, email, cpfecnpj, cidade, bairro, rua, numero);
        }
        else
        {
            Application ex = new Application();
            ex.DisplayAlerts = false;
            ex.Workbooks.Open(@"C:\Users\40809588897\Desktop\Programar\Semana 5\sistema_concessionaria\clientes.xls");
            int contador = 1;
            do
            {
                contador += 1;

            } while (ex.Cells[contador,1].Value != null);
            
            ex.Cells[contador,1].Value = nome;
            ex.Cells[contador,2].Value = email;
            ex.Cells[contador,3].Value = cpfecnpj;
            ex.Cells[contador,4].Value = cidade;
            ex.Cells[contador,5].Value = bairro;
            ex.Cells[contador,6].Value = rua;
            ex.Cells[contador,6].Value = numero;
            ex.ActiveWorkbook.Save();
            ex.Quit();
            ex.Dispose();
        }
    }
    public void Criarexcel(string nome, string email, string cpfecnpj, string cidade, string bairro, string rua, string numero)
    {    
        Application ex = new Application();
        ex.Workbooks.Add();
        ex.Cells[1,1].Value = nome;
        ex.Cells[1,2].Value = email;
        ex.Cells[1,3].Value = cpfecnpj;
        ex.Cells[1,4].Value = cidade;
        ex.Cells[1,5].Value = bairro;
        ex.Cells[1,6].Value = rua;
        ex.Cells[1,6].Value = numero;

        ex.ActiveWorkbook.SaveAs(@"C:\Users\40809588897\Desktop\Programar\Semana 5\sistema_concessionaria\clientes.xls");
        ex.Quit();
        ex.Dispose();
    }
}     
}      
    
          