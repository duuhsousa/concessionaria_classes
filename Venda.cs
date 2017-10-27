using System;
using NetOffice.ExcelApi;

namespace concessionaria_classes
{
    public class Venda
    {
        static Application ex = new Application();
        static void VendasDia(){
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
    }
}