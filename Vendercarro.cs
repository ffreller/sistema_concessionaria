using System;
using System.IO;
using NetOffice.ExcelApi;

namespace sistema_concessionaria{
    
    public class VenderCarro
    {
        public void Vendercarro()
        {
            Application ax = new Application();
            int contador1 = 1;
            int contador = 1;
            int preco1 = 1;
            string preco = "";
            Console.WriteLine("Digite seu CPF/CNPJ");
            string cpfcnpj = Console.ReadLine();
            bool vf = false;
            bool fv = false;
            ax.Workbooks.Open(@"C:\Users\40809588897\Desktop\Programar\Semana 5\sistema_concessionaria\clientes.xls");
            do
            {
                
                if(ax.Cells[contador1,3].Value.ToString() == cpfcnpj)
                {
                    vf = true;
                    break; 
                }
                contador1 += 1;
            }
            while (ax.Cells[contador1,3].Value != null);
                 
            if(vf==true)
            {
                Console.WriteLine("Carros Disponíveis:");
                Application ex = new Application();            
                ex.Workbooks.Open(@"C:\Users\40809588897\Desktop\Programar\Semana 5\sistema_concessionaria\carros.xls");
                do
                {
                    if(ex.Cells[contador,8].Value == null)
                    {
                        Console.WriteLine(ex.Cells[contador,1].Value.ToString() + "; " + ex.Cells[contador,2].Value.ToString() + "; " + ex.Cells[contador,3].Value.ToString());
                        string opcional1 = "";
                        string opcional2 = "";
                        string opcional3 = "";
                        if (ex.Cells[contador,4].Value.ToString() == "s")
                        {
                            opcional1 = "Ar-condicionado";
                        }
                        else{opcional1 = "Sem Ar-condicionado";}
                        if (ex.Cells[contador,5].Value.ToString() == "s")
                        {
                            opcional2 = "Airbag";
                        }
                        else{opcional1 = "Sem Airbag";}
                        
                        if (ex.Cells[contador,4].Value.ToString() == "s")
                        {
                            opcional3 = "ABS";
                        }
                        else{opcional1 = "Sem freios ABS";}
                        Console.WriteLine(opcional1 + "; " + opcional2 + "; " + opcional3 + ".");
                        }
                    contador += 1;    
                }
                while (ex.Cells[contador,1].Value != null);
                Console.WriteLine("Digite o nome do carro que deseja");
                string carroescolhido = Console.ReadLine();
                contador = 1;
                do
                {
                    if(ex.Cells[contador,1].Value.ToString() == carroescolhido)
                    {
                        fv = true;
                        break;
                    }
                    
                    else{contador += 1;}
                }
                while (ex.Cells[contador,1].Value != null);
                string vistaprazo = "";
                if(fv == true)
                {          
                    string vendido = "vendido";
                    ex.Cells[contador,8].Value = vendido;              
                    Console.WriteLine("Você escolheu o carro: " + carroescolhido);

                    bool opcao1 = false;
                    bool opcao2 = false;
                    string parcelas = "1";
                    int parcelas1 = 1;
                    preco = Convert.ToString (ex.Cells[contador,3].Value);
                    preco1 = Convert.ToInt16 (preco);
                    Console.WriteLine(preco1);
                    
                    do
                    {   
                        Console.WriteLine("Como deseja pagar? (digite 1 para a vista com 5% de desconto e 2 para a prazo)");
                        vistaprazo = Console.ReadLine();
                        switch(vistaprazo)
                        {   
                            case "1":
                                opcao1 = true;
                                preco1 = preco1 * 95/100;
                                Console.WriteLine("O preço fica " + preco1);
                                break;
                            case "2":
                                opcao1 = true;
                                Console.WriteLine("Em quantas parcelas deseja pagar?");
                                do
                                {
                                    Console.WriteLine("2, 4 ou 8 parcelas?");
                                    parcelas = Console.ReadLine();
                                    switch(parcelas)
                                    {
                                        case "2":
                                            opcao2 = true;
                                            break;
                                        case "4":
                                            opcao2 = true;
                                            break;
                                        case "8":
                                            opcao2 = true;
                                            break;
                                        default:
                                            Console.WriteLine("Opção Inválida.");
                                            break;
                                    }
                                    
                                }
                                while (opcao2 == false);
                                Console.WriteLine("O preço fica: ");
                                parcelas1 = Convert.ToInt16(parcelas);
                                int precoparcela = preco1/parcelas1; 
                                Console.WriteLine(parcelas + " parcelas de " + precoparcela + " reais");
                                break;
                        default:
                            Console.WriteLine("Opção Inválida.");
                        break;
                        }
                    }    
                    while(opcao1==false);
                
                    string cl1 = ax.Cells[contador1,1].Value.ToString();
                    string cl2 = ax.Cells[contador1,2].Value.ToString();
                    string cl3 = ax.Cells[contador1,3].Value.ToString();
                    string cl4 = ax.Cells[contador1,4].Value.ToString();
                    string cl5 = ax.Cells[contador1,5].Value.ToString();
                    string cl6 = ax.Cells[contador1,6].Value.ToString();
                    string cr1 = ex.Cells[contador,1].Value.ToString();
                    string cr2 = ex.Cells[contador,2].Value.ToString();
                    string cr3 = ex.Cells[contador,3].Value.ToString();
                    string cr4 = ex.Cells[contador,4].Value.ToString();
                    string cr5 = ex.Cells[contador,5].Value.ToString();
                    string cr6 = ex.Cells[contador,6].Value.ToString();
                
                
                
                    if(!File.Exists(@"C:\Users\40809588897\Desktop\Programar\Semana 5\sistema_concessionaria\vendas.xls"))
                    {
                        Criarexcelvenda(preco1, parcelas, cl1, cl2, cl3, cl4, cl5, cl6, cr1, cr2, cr3, cr4, cr5, cr6);
                        ex.ActiveWorkbook.Save();
                        ex.Quit();
                        ex.Dispose();
                    }
                    else
                    {
                        Application ox = new Application();
                        ox.DisplayAlerts = false;
                        ox.Workbooks.Open(@"C:\Users\40809588897\Desktop\Programar\Semana 5\sistema_concessionaria\vendas.xls");
                        int contador3 = 1;
                        do
                        {
                            contador3 += 1;

                        } while (ox.Cells[contador3,1].Value != null);
                        
                        
                        ox.Cells[contador3,1].Value = ax.Cells[contador1,1].Value;
                        ox.Cells[contador3,2].Value = ax.Cells[contador1,2].Value;
                        ox.Cells[contador3,3].Value = ax.Cells[contador1,3].Value;
                        ox.Cells[contador3,4].Value = ax.Cells[contador1,4].Value;
                        ox.Cells[contador3,5].Value = ax.Cells[contador1,5].Value;
                        ox.Cells[contador3,6].Value = ax.Cells[contador1,6].Value;
                        ox.Cells[contador3,7].Value = ex.Cells[contador,1].Value;
                        ox.Cells[contador3,8].Value = ex.Cells[contador,2].Value;
                        ox.Cells[contador3,9].Value = ex.Cells[contador,3].Value;
                        ox.Cells[contador3,10].Value = ex.Cells[contador,4].Value;
                        ox.Cells[contador3,11].Value = ex.Cells[contador,5].Value;
                        ox.Cells[contador3,12].Value = ex.Cells[contador,6].Value;
                        ox.Cells[contador3,13].Value = preco;
                        ox.Cells[contador3,14].Value = parcelas;
                    
                        ox.ActiveWorkbook.Save();
                        ox.Quit();
                        ox.Dispose();
                        ax.ActiveWorkbook.Save();
                        ax.Quit();
                        ax.Dispose();
                        ex.ActiveWorkbook.Save();
                        ex.Quit();
                        ex.Dispose();
                    }
                }
            else{
                CadastrarCarro carro1 = new CadastrarCarro();
                carro1.Cadastrarcarro();
                ax.ActiveWorkbook.Save();
                ax.Quit();
                ax.Dispose();
                ex.ActiveWorkbook.Save();
                ex.Quit();
                ex.Dispose();
                 
                } 
            }    
        else{
                CadastrarCliente cliente1 = new CadastrarCliente();
                cliente1.Cadastrarcliente();   
                ax.ActiveWorkbook.Save();
                ax.Quit();
                ax.Dispose();
                        
            }
        }
        
        
    public void Criarexcelvenda(int preco1, string parcelas, string cl1, string cl2, string cl3, string cl4, string cl5, string cl6, string cr1, string cr2, string cr3, string cr4, string cr5, string cr6)
    {    
        Application ix = new Application();
        ix.Workbooks.Add();
        ix.Cells[1,1].Value = cl1;
        ix.Cells[1,2].Value = cl2;
        ix.Cells[1,3].Value = cl3;
        ix.Cells[1,4].Value = cl4;
        ix.Cells[1,5].Value = cl5;
        ix.Cells[1,6].Value = cl6;
        ix.Cells[1,7].Value = cr1;
        ix.Cells[1,8].Value = cr2;
        ix.Cells[1,9].Value = cr3;
        ix.Cells[1,10].Value = cr4;
        ix.Cells[1,11].Value = cr5;
        ix.Cells[1,12].Value = cr6;
        ix.Cells[1,13].Value = preco1;
        ix.Cells[1,14].Value = parcelas;
        
        ix.ActiveWorkbook.SaveAs(@"C:\Users\40809588897\Desktop\Programar\Semana 5\sistema_concessionaria\vendas.xls");
        ix.Quit();
        ix.Dispose();
    }
            
    


        }
    }

    
    