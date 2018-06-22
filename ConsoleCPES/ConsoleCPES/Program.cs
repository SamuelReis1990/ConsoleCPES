using System;
using System.Diagnostics;
using System.Linq;
using CPES.Models;

namespace ConsoleCPES
{
    class Program
    {
        static void Main(string[] args)
        {
            Erro erro = new Erro();

            try
            {

                Console.WriteLine("");
                Console.SetCursorPosition((Console.WindowWidth - "Bem vindo ao CPES".Length) / 2, Console.CursorTop);
                Console.WriteLine("Bem vindo ao CPES");
#if true

                string caminhoArquivoParametro = "C:\\Users\\samue\\Desktop\\Planilha telegrama reembolso.xls";
                string caminhoArquivoEndereco = "C:\\Users\\samue\\Desktop\\RJCDC04.xls";
                string caminhoArquivoLotacao = "C:\\Users\\samue\\Desktop\\RJCD04.xls";
                string tabelaParametro = "Sheet1$";
                string tabelaEndereco = "RJCDC04$";
                string tabelaLotacao = "RJCD04$";
                int numColMatricula = 1;
                int numColMatriculaEndereco = 1;
                int numColMatriculaLotacao = 1;
                int numColNome = 2;
                int numColEndereco = 5;
                int numColNumero = 6;
                int numColComplemento = 7;
                int numColBairro = 8;
                int numColCidade = 9;
                int numColUF = 10;
                int numColCEP = 11;
                int numColLotacao = 26;

#else

                #region Entrada arquivo de parametro

                Console.WriteLine("");
                Console.SetCursorPosition((Console.WindowWidth - "1 ARQUIVO".Length) / 2, Console.CursorTop);
                Console.WriteLine("1 ARQUIVO");

                Console.WriteLine("\nInforme o arquivo que contém as matrículas a serem listadas*:");
                Console.WriteLine("Dica: Precisa informar o local onde se encontra o arquivo seguido do");
                Console.WriteLine("nome completo do arquivo e sua extensão.");
                Console.WriteLine(@"Ex: C:\User\Desktop\Planilha telegrama reembolso.xls");
                Console.WriteLine("Obs: O campo é obrigatório e portanto deve ser informado!");
                Console.WriteLine("IMPORTANTE: O arquivo deve ser uma planilha excel com a extensão em .XLS");
                Console.WriteLine("Digite abaixo ou copie e cole o local onde se encontra o arquivo e pressione a tecla Enter:");

                string caminhoArquivoParametro;
                while (true)
                {
                    if (!File.Exists(caminhoArquivoParametro = Console.ReadLine()))
                    {
                        Console.WriteLine("\nArquivo não encontrado!");
                        Console.WriteLine("Dica: Verifique se digitou/colou corretamente o local ou o nome do arquivo desejado.");
                        Console.WriteLine("Tente novamente:");
                    }
                    else
                    {
                        break;
                    }
                }

                Console.WriteLine("\nInforme o nome da tabela que contém as matrículas*:");
                Console.WriteLine("Ex: Sheet1/Plan1");
                Console.WriteLine("Obs: O campo é obrigatório e portanto deve ser informado!");
                Console.WriteLine("Digite abaixo o nome da tabela e pressione a tecla Enter:");

                string tabelaParametro = Console.ReadLine() + "$";
                while (true)
                {
                    if (!String.IsNullOrEmpty(tabelaParametro) && tabelaParametro != "$")
                    {
                        break;
                    }
                    else
                    {
                        Console.WriteLine("\nO nome da tabela é obrigatório!");
                        Console.WriteLine("Por favor informe o nome da tabela:");
                        tabelaParametro = Console.ReadLine() + "$";
                    }
                }

                Console.WriteLine("\nInforme o número da coluna que contém as matrículas*:");
                Console.WriteLine("Ex: Coluna A, seria o número 1 e assim sucessivamente");
                Console.WriteLine("Obs: O campo é obrigatório e portanto deve ser preenchido!");
                Console.WriteLine("Digite abaixo o número da coluna e pressione a tecla Enter:");

                int numColMatricula;
                while (true)
                {
                    if (!int.TryParse(Console.ReadLine(), out numColMatricula))
                    {
                        Console.WriteLine("\nInforme somente números!");
                        Console.WriteLine("Tente novamente:");
                    }
                    else
                    {
                        break;
                    }
                }

                #endregion

                #region Entrada arquivo de endereço

                Console.WriteLine("");
                Console.SetCursorPosition((Console.WindowWidth - "2 ARQUIVO".Length) / 2, Console.CursorTop);
                Console.WriteLine("2 ARQUIVO");

                Console.WriteLine("\nInforme o arquivo que contém os endereços a serem recuperados*:");
                Console.WriteLine("Dica: Precisa informar o local onde se encontra o arquivo seguido do");
                Console.WriteLine("nome completo do arquivo e sua extensão.");
                Console.WriteLine(@"Ex: C:\User\Desktop\RJCDC04.xls");
                Console.WriteLine("Obs: O campo é obrigatório e portanto deve ser informado!");
                Console.WriteLine("IMPORTANTE: O arquivo deve ser uma planilha excel com a extensão em .XLS");
                Console.WriteLine("Digite abaixo ou copie e cole o local onde se encontra o arquivo e pressione a tecla Enter:");

                string caminhoArquivoEndereco;
                while (true)
                {
                    if (!File.Exists(caminhoArquivoEndereco = Console.ReadLine()))
                    {
                        Console.WriteLine("\nArquivo não encontrado!");
                        Console.WriteLine("Dica: Verifique se digitou/colou corretamente o local ou o nome do arquivo desejado.");
                        Console.WriteLine("Tente novamente:");
                    }
                    else
                    {
                        break;
                    }
                }

                Console.WriteLine("\nInforme o nome da tabela que contém os endereços*:");
                Console.WriteLine("Ex: Sheet1/Plan1/RJCDC04");
                Console.WriteLine("Obs: O campo é obrigatório e portanto deve ser informado!");
                Console.WriteLine("Digite abaixo o nome da tabela e pressione a tecla Enter:");

                string tabelaEndereco = Console.ReadLine() + "$";
                while (true)
                {
                    if (!String.IsNullOrEmpty(tabelaEndereco) && tabelaEndereco != "$")
                    {
                        break;
                    }
                    else
                    {
                        Console.WriteLine("\nO nome da tabela é obrigatório!");
                        Console.WriteLine("Por favor informe o nome da tabela:");
                        tabelaEndereco = Console.ReadLine() + "$";
                    }
                }

                Console.WriteLine("\nInforme os números das colunas que contém os campos de endereços:");
                Console.WriteLine("Ex: Coluna A, seria o número 1 e assim sucessivamente");
                Console.WriteLine("Obs: Essas colunas são opcionais por tanto se não digitar nada será assumido");
                Console.WriteLine("um valor default!");
                Console.WriteLine("O valor default assumido, caso não digite NADA será:");
                Console.WriteLine("1, 5, 6, 7, 8, 9, 10, 11 - extamente nesta ordem e quantidade,");
                Console.WriteLine("esses valores seriam para cada coluna a saber:");
                Console.WriteLine("o valor 1 para a coluna da matrícula");
                Console.WriteLine("o valor 5 para a coluna do logradouro do endereço");
                Console.WriteLine("o valor 6 para a coluna do número do endereço");
                Console.WriteLine("o valor 7 para a coluna do complemento do endereço");
                Console.WriteLine("o valor 8 para a coluna do bairro do endereço");
                Console.WriteLine("o valor 9 para a coluna da cidade do endereço");
                Console.WriteLine("o valor 10 para a coluna da UF do endereço");
                Console.WriteLine("o valor 11 para a coluna do CEP do endereço");
                Console.WriteLine("IMPORTANTE: Caso os valores citados acima não sejam");
                Console.WriteLine("correspondentes as colunas do seu arquivo atual,");
                Console.WriteLine("então deve-se digitar os valores correspondentes das colunas");
                Console.WriteLine("seguindo a mesma ordem e quantidade acima.");
                Console.WriteLine("Digite abaixo os números de todas as colunas separados por virgula");
                Console.WriteLine("e pressione a tecla Enter:");

                int numColMatriculaEndereco = 1;
                int numColEndereco = 5;
                int numColNumero = 6;
                int numColComplemento = 7;
                int numColBairro = 8;
                int numColCidade = 9;
                int numColUF = 10;
                int numColCEP = 11;

                string[] colunasEndereco = Console.ReadLine().Trim().Split(',');
                if (colunasEndereco != null && colunasEndereco[0] != "")
                {
                    while (true)
                    {
                        if (colunasEndereco.Count() == 8)
                        {
                            while (true)
                            {
                                if (!int.TryParse(colunasEndereco[0], out numColMatriculaEndereco) || 
                                    !int.TryParse(colunasEndereco[1], out numColEndereco) ||
                                    !int.TryParse(colunasEndereco[2], out numColNumero) || 
                                    !int.TryParse(colunasEndereco[3], out numColComplemento) ||
                                    !int.TryParse(colunasEndereco[4], out numColBairro) || 
                                    !int.TryParse(colunasEndereco[5], out numColCidade) ||
                                    !int.TryParse(colunasEndereco[6], out numColUF) ||
                                    !int.TryParse(colunasEndereco[7], out numColCEP))
                                {
                                    Console.WriteLine("\nInforme somente números!");
                                    Console.WriteLine("Tente novamente:");
                                    colunasEndereco = Console.ReadLine().Trim().Split(',');
                                }
                                else
                                {
                                    break;
                                }
                            }

                            break;
                        }
                        else
                        {
                            Console.WriteLine("\nInforme oito números separados por virgulas correspondentes");
                            Console.WriteLine("as oito colunas somente, conforme exemplificado.");
                            Console.WriteLine("Tente novamente:");
                            colunasEndereco = Console.ReadLine().Trim().Split(',');
                        }
                    }
                }

                #endregion

                #region Entrada arquivo de lotação

                Console.WriteLine("");
                Console.SetCursorPosition((Console.WindowWidth - "3 ARQUIVO".Length) / 2, Console.CursorTop);
                Console.WriteLine("3 ARQUIVO");

                Console.WriteLine("\nInforme o arquivo que contém as lotações e nomes a serem recuperados*:");
                Console.WriteLine("Dica: Precisa informar o local onde se encontra o arquivo seguido do");
                Console.WriteLine("nome completo do arquivo e sua extensão.");
                Console.WriteLine(@"Ex: C:\User\Desktop\RJCD04.xls");
                Console.WriteLine("Obs: O campo é obrigatório e portanto deve ser informado!");
                Console.WriteLine("IMPORTANTE: O arquivo deve ser uma planilha excel com a extensão em .XLS");
                Console.WriteLine("Digite abaixo ou copie e cole o local onde se encontra o arquivo e pressione a tecla Enter:");

                string caminhoArquivoLotacao;
                while (true)
                {
                    if (!File.Exists(caminhoArquivoLotacao = Console.ReadLine()))
                    {
                        Console.WriteLine("\nArquivo não encontrado!");
                        Console.WriteLine("Dica: Verifique se digitou/colou corretamente o local ou o nome do arquivo desejado.");
                        Console.WriteLine("Tente novamente:");
                    }
                    else
                    {
                        break;
                    }
                }

                Console.WriteLine("\nInforme o nome da tabela que contém as lotações e nomes*:");
                Console.WriteLine("Ex: Sheet1/Plan1/RJCD04");
                Console.WriteLine("Obs: O campo é obrigatório e portanto deve ser informado!");
                Console.WriteLine("Digite abaixo o nome da tabela e pressione a tecla Enter:");

                string tabelaLotacao = Console.ReadLine() + "$";
                while (true)
                {
                    if (!String.IsNullOrEmpty(tabelaLotacao) && tabelaLotacao != "$")
                    {
                        break;
                    }
                    else
                    {
                        Console.WriteLine("\nO nome da tabela é obrigatório!");
                        Console.WriteLine("Por favor informe o nome da tabela:");
                        tabelaLotacao = Console.ReadLine() + "$";
                    }
                }

                Console.WriteLine("\nInforme os números das colunas que contém os campos de nome e lotação:");
                Console.WriteLine("Ex: Coluna A, seria o número 1 e assim sucessivamente");
                Console.WriteLine("Obs: Essas colunas são opcionais por tanto se não digitar nada será assumido");
                Console.WriteLine("um valor default!");
                Console.WriteLine("O valor default assumido, caso não digite NADA será:");
                Console.WriteLine("1, 2, 26 - extamente nesta ordem e quantidade,");
                Console.WriteLine("esses valores seriam para cada coluna a saber:");
                Console.WriteLine("o valor 1 para a coluna do matrícula");
                Console.WriteLine("o valor 2 para a coluna do nome");
                Console.WriteLine("o valor 26 para a coluna da lotação");
                Console.WriteLine("IMPORTANTE: Caso os valores citados acima não sejam");
                Console.WriteLine("correspondentes as colunas do seu arquivo atual,");
                Console.WriteLine("então deve-se digitar os valores correspondentes das colunas");
                Console.WriteLine("seguindo a mesma ordem e quantidade acima.");
                Console.WriteLine("Digite abaixo os números de todas as colunas separados por virgula");
                Console.WriteLine("e pressione a tecla Enter:");

                int numColMatriculaLotacao = 1;
                int numColNome = 2;
                int numColLotacao = 26;

                string[] colunasLotacao = Console.ReadLine().Trim().Split(',');
                if (colunasLotacao != null && colunasLotacao[0] != "")
                {
                    while (true)
                    {
                        if (colunasLotacao.Count() == 3)
                        {
                            while (true)
                            {
                                if (!int.TryParse(colunasLotacao[0], out numColMatriculaLotacao) || 
                                    !int.TryParse(colunasLotacao[1], out numColNome) ||
                                    !int.TryParse(colunasLotacao[2], out numColLotacao))
                                {
                                    Console.WriteLine("\nInforme somente números!");
                                    Console.WriteLine("Tente novamente:");
                                    colunasLotacao = Console.ReadLine().Trim().Split(',');
                                }
                                else
                                {
                                    break;
                                }
                            }

                            break;
                        }
                        else
                        {
                            Console.WriteLine("\nInforme três números separados por virgulas correspondentes");
                            Console.WriteLine("as três colunas somente, conforme exemplificado.");
                            Console.WriteLine("Tente novamente:");
                            colunasLotacao = Console.ReadLine().Trim().Split(',');
                        }
                    }
                }

                #endregion

                Console.WriteLine("");
                Console.SetCursorPosition((Console.WindowWidth - "AGUARDE...".Length) / 2, Console.CursorTop);
                Console.WriteLine("AGUARDE...");                
#endif

                #region Grava os parametros de entrada caso haja algum tipo de erro

                erro.CaminhoArquivoParametro = caminhoArquivoParametro;
                erro.CaminhoArquivoEndereco = caminhoArquivoEndereco;
                erro.CaminhoArquivoLotacao = caminhoArquivoLotacao;
                erro.TabelaParametro = tabelaParametro;
                erro.TabelaEndereco = tabelaEndereco;
                erro.TabelaLotacao = tabelaLotacao;
                erro.NumColMatricula = numColMatricula;
                erro.NumColMatriculaEndereco = numColMatriculaEndereco;
                erro.NumColMatriculaLotacao = numColMatriculaLotacao;
                erro.NumColNome = numColNome;
                erro.NumColEndereco = numColEndereco;
                erro.NumColNumero = numColNumero;
                erro.NumColComplemento = numColComplemento;
                erro.NumColBairro = numColBairro;
                erro.NumColCidade = numColCidade;
                erro.NumColUF = numColUF;
                erro.NumColCEP = numColCEP;
                erro.NumColLotacao = numColLotacao;

                #endregion                

                var planilhaParametro = CPES.CPES.GetExcel(caminhoArquivoParametro, tabelaParametro, numColMatricula, numColMatricula, numColMatricula);
                var planilhaEndereco = CPES.CPES.GetExcel(caminhoArquivoEndereco, tabelaEndereco, numColMatriculaEndereco, numColNumero, numColCEP);
                var planilhaLotacao = CPES.CPES.GetExcel(caminhoArquivoLotacao, tabelaLotacao, numColMatriculaLotacao, numColMatricula, numColMatricula);

                var colParametro = planilhaParametro.Select(s => s.Table.Columns).First();
                var colEndereco = planilhaEndereco.Select(s => s.Table.Columns).First();
                var colLotacao = planilhaLotacao.Select(s => s.Table.Columns).First();

                string[] retorno = CPES.CPES.GetDados(planilhaParametro, planilhaEndereco, planilhaLotacao, colParametro, colEndereco, colLotacao
                                            , numColMatricula
                                            , numColMatriculaEndereco
                                            , numColMatriculaLotacao
                                            , numColNome
                                            , numColEndereco
                                            , numColNumero
                                            , numColComplemento
                                            , numColBairro
                                            , numColCidade
                                            , numColUF
                                            , numColCEP
                                            , numColLotacao
                                           );

                Console.WriteLine("\nOperação realizada com sucesso!");
                Console.WriteLine(string.Format("Foi gerado o arquivo {0}", retorno[0]));
                Console.WriteLine(string.Format("em {0}", retorno[1]));
                Console.WriteLine(string.Format("com um total de {0} registros.", retorno[2]));
                Console.WriteLine("Aperte qualquer tecla para finalizar...");
                Console.ReadKey();
            }
            catch (Exception e)
            {
                var stackTrace = new StackTrace(e, true);
                var frames = stackTrace.GetFrames();

                string textoStackTrace = string.Empty;
                foreach (var frame in frames)
                {
                    textoStackTrace += "Metodo do erro: " + frame.GetMethod().Name.ToString() + " Linha do erro: " + frame.GetFileLineNumber().ToString() + ", ";
                }

                erro.Mensagem = e.Message;
                erro.StackTrace = textoStackTrace;
                string[] retorno = CPES.CPES.CreateCsvErro(erro);
                Console.WriteLine("\nHouve um erro!");
                Console.WriteLine(string.Format("Foi gerado um arquivo de log de erro {0}", retorno[0]));
                Console.WriteLine(string.Format("em {0}", retorno[1]));
                Console.WriteLine("Tente novamente mais tarde, se o erro permanecer envie o arquivo de log gerado");
                Console.WriteLine("ao administrador do sistema.");
                Console.WriteLine("Aperte qualquer tecla para finalizar...");
                Console.ReadKey();
            }
        }
    }
}
