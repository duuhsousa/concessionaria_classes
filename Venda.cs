using System;
using System.IO;
using NetOffice.ExcelApi;
using System.Text.RegularExpressions;

namespace concessionaria_classes
{
    
    public class Venda
    {
        Application ex = new Application();
        Regex rgx = new Regex(@"^\S{3}\d{4}$");
        Cliente cliente = new Cliente();
        string formaPagamento;

            public void VendasDia(){

            string op1,data;



            do{

                Console.Write("Digite o dia para pesquisa (DD/MM/AAAA): ");

                data = Console.ReadLine();

                ex.Workbooks.Open(@"C:\Concessionaria\Cadastro_Cliente.xls");

                

                int cont=1;

                do{

                    if(ex.Cells[cont,3].Value.ToString() == data){

                        Console.WriteLine(ex.Cells[cont,1].Value.ToString()+"\t");

                        Console.WriteLine(ex.Cells[cont,2].Value.ToString()+"\t");

                        Console.WriteLine(ex.Cells[cont,3].Value.ToString()+"\t");

                        Console.WriteLine(ex.Cells[cont,4].Value.ToString()+"\t");

                        if(ex.Cells[cont,4].Value.ToString() == "VISTA"){

                            Console.WriteLine(ex.Cells[cont,6].Value.ToString()+"\n");

                        }else{

                            Console.WriteLine(int.Parse(ex.Cells[cont,6].Value.ToString())*int.Parse(ex.Cells[cont,5].Value.ToString())+"\t");

                        }

                    }

                    cont++;

                }while(ex.Cells[cont,1].Value!=null);







            do{

                    Console.Write("\nDeseja realizar um novo cadastro? (S ou N)");

                    op1 = Console.ReadLine();

                } while (op1!="S" && op1!="N" && op1!="s" && op1!="n");

            } while(op1=="S" || op1=="s");

        }

        public void RealizarVendas(){
            int cont=1;
            string placa;

            ex.Workbooks.Open(@"C:\Concessionaria\Cadastro_Carro.xls");

            do{
                    if(ex.Cells[cont,7].Value.ToString() == "0"){
                        
                        Console.Write(ex.Cells[cont,1].Value+"\t");
                        Console.Write(ex.Cells[cont,2].Value+"-");
                        Console.Write(ex.Cells[cont,3].Value+"\t");
                        Console.Write(ex.Cells[cont,4].Value+@"\");
                        Console.Write(ex.Cells[cont,5].Value+"\t");
                        Console.Write("R$"+ex.Cells[cont,6].Value+"\n");

                    }
                cont++;
            }while(ex.Cells[cont,1].Value!=null);

            do{
            Console.Write("\n\nInforme placa do carro a ser comprado: ");    
            placa = Console.ReadLine();
            }while(!rgx.IsMatch(placa));

            cont=1;

            //Validar se já foi comprado
            do{
                    if(ex.Cells[cont,1].Value.ToString() == placa && ex.Cells[cont,7].Value.ToString()!="0"){        
                        Console.WriteLine("Carro já foi comprado!");
                    }
                cont++;
            }while(ex.Cells[cont,1].Value!=null);

            //Voltando para posição da placa
            cont--;

            do{            
            Console.WriteLine("Forma de pagamento: (VISTA / PRAZO) ");
            formaPagamento = Console.ReadLine();
            }while(formaPagamento.ToUpper()!="VISTA" && formaPagamento.ToUpper()!="PRAZO");

            int parcelas=0;
            double valor = 0;

            //Voltar a posição encontrada
            cont=cont-1;

            if(formaPagamento.ToUpper()=="VISTA"){
                Console.WriteLine("cont"+cont);
                Console.WriteLine(ex.Cells[cont,6].Value.ToString());
                Console.WriteLine(double.Parse(ex.Cells[cont,6].Value.ToString()).ToString());
                valor = double.Parse(ex.Cells[cont,6].Value.ToString())-((double.Parse(ex.Cells[cont,6].Value.ToString())*5)/100);
            }
            else
            {
                Console.Write("Quantidade de Parcelas: ");
                parcelas = int.Parse(Console.ReadLine());
                valor = double.Parse(ex.Cells[cont,6].Value.ToString())/parcelas;
            }

            ex.Quit();

            Console.WriteLine("Informe documento do cliente: ");           
            string doc = Console.ReadLine();
            int linha = cliente.PesquisaDocumento(doc); 

            if(linha==0){
                Console.WriteLine("Cliente não cadastrado!");  
                //cliente.CadastrarClientes();
                //validar
            }
            
            ex.Workbooks.Open(@"C:\Concessionaria\Cadastro_Venda.xls");

            ex.Cells[linha,1].Value = doc;
            ex.Cells[linha,2].Value = placa;
            ex.Cells[linha,3].Value = DateTime.Now;
            ex.Cells[linha,4].Value = formaPagamento;
            ex.Cells[linha,5].Value = parcelas;
            ex.Cells[linha,6].Value = valor;

            ex.ActiveWorkbook.Save();
            ex.Quit();
            ex.Dispose();

            Console.WriteLine("Venda realizada!");
        }
    }
}