using System;
using System.IO;
using NetOffice.ExcelApi;

namespace concessionaria_classes
{
    public class Cliente
    {
        private string nome {get;set;}
        private string doc {get;set;}
        private string email {get;set;}
        private string tipo {get;set;}
        public Validacao validacao = new Validacao();
        public Endereco endereco = new Endereco();
 
        public void CadastrarClientes()
        {
            string op1;
            int valid = 0;
            int duplicado;

            do{
                Console.WriteLine("\nCADASTRO DE CLIENTES: \n");
                do{
                    Console.Write("Digite 1 para CPF e 2 para CNPJ: ");
                    tipo = Console.ReadLine();
                }while(tipo!="1" && tipo!="2");
                do{
                    if(tipo=="1"){ 
                        do{
                            Console.Write("CPF: ");
                            doc = Console.ReadLine();
                            duplicado = PesquisaDocumento(doc); 
                            
                            if(doc.Length!=11){
                                Console.WriteLine("Formato de CPF inválido!");
                            }

                        }while(doc.Length!=11 || duplicado!=0);
                        valid = validacao.ValidarCPF(doc);
                    }
                    else{
                        do{
                            Console.Write("CNPJ: ");
                            doc = Console.ReadLine();    
                            duplicado = PesquisaDocumento(doc); 

                            if(doc.Length!=14){
                                Console.WriteLine("Formato de CNPJ inválido!");
                            }

                        }while(doc.Length!=14 || duplicado!=0);
                        valid = validacao.ValidarCNPJ(doc);
                    }
                }while(valid!=1);

                Application ex = new Application();
                ex.Workbooks.Add();
                Console.Write("Nome completo: ");
                ex.Cells[1,1].Value = Console.ReadLine();
                Console.Write("Email: ");
                ex.Cells[1,2].Value = Console.ReadLine();
                ex.ActiveWorkbook.SaveAs("Cadstro_Clientes.xls");
                ex.Quit();
                do
                {
                    Console.Write("\nDeseja realizar um novo cadastro? (S ou N)");
                    op1 = Console.ReadLine();
                } while (op1!="S" && op1!="N" && op1!="s" && op1!="n");
            } while(op1=="S" || op1=="s");
        }

         public int PesquisaDocumento(string docCliente)
        {
            if(File.Exists("cliente.csv")){
                String[] clientes = File.ReadAllLines("cliente.csv");
                String[] dadosCliente;

                foreach(string cliente in clientes){
                    dadosCliente = cliente.Split(';');
                    if(dadosCliente[0].Equals(docCliente)){
                        Console.WriteLine("\nCliente já cadastrado no sistema!\n");
                        return 1;
                    }
                }
            }
            return 0;
        }

    }
}