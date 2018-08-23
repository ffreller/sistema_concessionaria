using System;
using System.IO;

namespace sistema_concessionaria
{class Program

    {static void Main(string[] args)
        {string opcao1 = "";

        do
            
            {Console.WriteLine("Digite a opção");
            Console.WriteLine("1 - Cadastrar cliente");
            Console.WriteLine("2 - Cadastrar carro");
            Console.WriteLine("3 - Vender carro");
            Console.WriteLine("4 - Listar carros vendidos");
            Console.WriteLine("5 - Sair");
            opcao1 = Console.ReadLine();
            
            switch(opcao1)
            
                {case "1": CadastrarCliente cliente1 = new CadastrarCliente();
                         cliente1.Cadastrarcliente(); 
                break;
                case "2":CadastrarCarro carro1 = new CadastrarCarro();
                         carro1.Cadastrarcarro();   
                break;
                case "3":VenderCarro venda1 = new VenderCarro();
                         venda1.Vendercarro();   
                break;
                case "4": ListarCarros lista1 = new ListarCarros();
                        lista1.Listarcarros(); 
                break;
                case "5":
                        {Console.WriteLine("Deseja realmente sair(s ou n)");
                        string sair = Console.ReadLine();
                        if(sair.ToLower().Contains("s"))
                            Environment.Exit(0);
                        else if(!sair.ToLower().Contains("n"))    
                            {opcao1 = "0";
                            Console.WriteLine("Opção Inválida");
                            }
                        else
                            {opcao1 = "0";
                            }
                        }                 
                             
                break;
            }
        }
        while (opcao1 != "5");
        }
    } 
}       