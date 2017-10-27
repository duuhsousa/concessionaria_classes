using System;
using NetOffice.ExcelApi;

namespace concessionaria_classes
{
    public class Carro
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
                        //do{
                            Console.Write("CPF: ");
                            doc = Console.ReadLine();
                            //duplicado = PesquisaDocumento(doc); 
                            
                            if(doc.Length!=11){
                                Console.WriteLine("Formato de CPF inválido!");
                            }

                        //}while(doc.Length!=11 || duplicado!=0);
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

                ex.Workbooks.Open(@"C:\Concessionaria\Cadastro_Cliente.xls");

                int cont=1;
                
                do{
                    cont++;
                }while(ex.Cells[cont,1].Value!=null);
                
                
                ex.Cells[cont,1].Value = doc;
                Console.Write("Nome: ");
                ex.Cells[cont,2].Value = Console.ReadLine();
                Console.Write("Endereço: ");
                ex.Cells[cont,3].Value = Console.ReadLine();
                Console.Write("Cidade: ");
                ex.Cells[cont,4].Value = Console.ReadLine();
                Console.Write("Estado: ");
                ex.Cells[cont,5].Value = Console.ReadLine();
                Console.Write("CEP: ");
                ex.Cells[cont,6].Value = Console.ReadLine();

                ex.ActiveWorkbook.Save();
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
            if(File.Exists(@"C:\Concessionaria\Cadastro_Clientes.xls")){
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