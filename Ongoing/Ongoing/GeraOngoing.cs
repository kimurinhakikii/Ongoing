using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.IO;



using System.Xml.Linq;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Diagrams;
//using DocumentFormat.OpenXml.Office2010.Excel;
//using DocumentFormat.OpenXml.Wordprocessing;


namespace Ongoing
{
    public partial class GeraOngoing : Form
    {
        public GeraOngoing()
        {

            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Execucao();
            MessageBox.Show("Execução Concluída");
        }

        private void Execucao()
        {
            try
            {
                textBox1.Text = DateTime.Now.ToString();

                // Cria a pasta caso não exista
                if (!Directory.Exists(folderPathImagemPPs))
                {
                    Directory.CreateDirectory(folderPathImagemPPs);
                }

                if (!Directory.Exists(folderPathImagemEstruturante))
                {
                    Directory.CreateDirectory(folderPathImagemEstruturante);
                }

                if (!Directory.Exists(folderPathImagemExterno))
                {
                    Directory.CreateDirectory(folderPathImagemExterno);
                }

                if (!Directory.Exists(folderPathImagemEletronico))
                {
                    Directory.CreateDirectory(folderPathImagemEletronico);
                }

                File.Delete(caminhoFichaGerada);
                File.Copy(caminhoFechaLimpa, caminhoFichaGerada);

                geraPP();
                geraEstruturante();
                geraEstruturanteInfra();
                geraProjetosExternos();
                geraProjetosEletronicos();
                deleteFiles();
                textBox2.Text = DateTime.Now.ToString();
            }
            catch (Exception ex) { }
        }

        public static void deleteFiles()
        {
            try
            {
                Directory.Delete(folderPathImagemPPs, true);
                Directory.Delete(folderPathImagemEstruturante, true);
                Directory.Delete(folderPathImagemExterno, true);
                Directory.Delete(folderPathImagemEletronico, true);
            }
            catch { }

        }


        private static string caminhoFechaLimpa = @"C:\\Users\\80105812.REDECORP\\Desktop\\PowerPoint\\Ongoing Projetos de Segurança Corporativa.pptx";
        private static string caminhoFichaGerada = @"C:\Temp_Ongoing\Ongoing Projetos de Segurança Corporativa.pptx";
        private static string folderPathImagemPPs = @"C:\Temp_Ongoing\PPs";
        private static string folderPathImagemEstruturante = @"C:\Temp_Ongoing\Estruturante";
       private static string folderPathImagemExterno = @"C:\Temp_Ongoing\ProjetosExternos";
        private static string folderPathImagemEletronico = @"C:\Temp_Ongoing\ProjetosEletronicos";

        private void geraPP()
        {
            try
            {
                List<string> valoresFotos = new List<string>();
                string sqlConsult = "SELECT T.Id, NomePP, GerenciaDemandante, Solicitante, DataAtualizacao, Responsavel, DataCadatrado, ContentType, Bytes, NomeArquivo" + Environment.NewLine
                                    + " FROM ( " + Environment.NewLine
                                    + " SELECT   P.[Id],[NomePP],  area.NomeAreaGestora as GerenciaDemandante, S.NomeUsuario as 'Solicitante',CONVERT(varchar(10), P.DataAtualizacao, 103) as 'DataAtualizacao', " + Environment.NewLine
                                    + " u.NomeUsuario as 'Responsavel' , B.DataCadatrado " + Environment.NewLine
                                    + " FROM [PORTAL_CORPORATIVO_PRD].[dbo].[PPs] P " + Environment.NewLine
                                    + " inner join [dbo].[Usuarios] U on P.AnalistaResponsavelId = U.UsuarioId " + Environment.NewLine
                                    + " left join  [dbo].[Usuarios] S on P.SolicitanteId = S.UsuarioId " + Environment.NewLine
                                    + " left join [dbo].[AreasGestoras] area On S.AreaGestoraId = area.Id " + Environment.NewLine
                                    + " left join " + Environment.NewLine
                                    + " ( " + Environment.NewLine
                                    + "   SELECT [ProjetoId], Max([Inicio]) DataCadatrado " + Environment.NewLine
                                    + "   FROM [PORTAL_CORPORATIVO_PRD].[dbo].[ArquivosBiWeekly] A " + Environment.NewLine
                                    + "   where [TipoProjeto] = 1 " + Environment.NewLine
                                    + "   group by  [ProjetoId] " + Environment.NewLine
                                    + " ) B on P.Id = B.ProjetoId " + Environment.NewLine
                                    + " where EtapaPPId in (6,7,8,9) AND Deleted = 0 and P.id != 3110 " + Environment.NewLine
                                    + " group by  P.[Id],[NomePP], u.NomeUsuario,B.DataCadatrado, p.DataAberturaPSS, p.DataCadastro, S.NomeUsuario, area.NomeAreaGestora, P.DataAtualizacao " + Environment.NewLine
                                    + "   )T " + Environment.NewLine
                                    + "   inner join [PORTAL_CORPORATIVO_PRD].[dbo].[ArquivosBiWeekly] Arquivo on Arquivo.ProjetoId = T.Id and Arquivo.Inicio = T.DataCadatrado " + Environment.NewLine
                                    + "   ORDER BY T.Id ASC";

                SqlCommand cmd = Util.sqlReader(sqlConsult);
                SqlDataReader rdr = cmd.ExecuteReader();

                if (rdr.HasRows)
                {
                    while (rdr.Read())
                    {
                        var valoresLinha = new List<string> { rdr["NomePP"].ToString(), rdr["GerenciaDemandante"].ToString(), rdr["Solicitante"].ToString(), rdr["DataAtualizacao"].ToString() };
                        if (rdr["NomeArquivo"].ToString() != null)
                        {
                            byte[] imagemBytes = (byte[])rdr["Bytes"];
                            string nomeOriginal = rdr["Id"].ToString() + "_" + rdr["NomeArquivo"].ToString();
                            string extensao = System.IO.Path.GetExtension(nomeOriginal);

                            valoresFotos.Add(rdr["NomePP"].ToString() + ";" + nomeOriginal);

                            // Define nome do arquivo salvo
                            string caminhoCompleto = System.IO.Path.Combine(folderPathImagemPPs, nomeOriginal);

                            // Salva a imagem
                            File.WriteAllBytes(caminhoCompleto, imagemBytes);
                        }
                        AdicionarLinha(caminhoFichaGerada, valoresLinha, "PPs EM ANDAMENTO", 0);
                    }
                    rdr.Close();
                    cmd.Connection.Close();

                    var idPagina = ObterPosicaoDoSlide(caminhoFichaGerada, "PPs EM ANDAMENTO");
                    for (int i = 0; valoresFotos.Count > 0; i++)
                    {
                        var Photo = valoresFotos[i].Split(';');
                        AdicionarImagem.InserirSlideComImagemNaOrdem(caminhoFichaGerada, folderPathImagemPPs + @"\" + Photo[1], Photo[0], idPagina + (i + 1));
                    }
                }

            }
            catch (Exception ex) { }
        }


        private void geraEstruturante()
        {
            try
            {
                List<string> valoresFotos = new List<string>();
                string sqlConsult = "SELECT T.Id, T.NomeProjeto, T.TipoIdentificador, T.IdPSS, T.GerenciaDemandante, T.Solicitante, T.DataAtualizacao, T.Responsavel, T.DataCadatrado, ContentType, Bytes, NomeArquivo " + Environment.NewLine
                                    + "FROM (" + Environment.NewLine
                                    + "        SELECT  P.[Id], " + Environment.NewLine
                                    + "        CASE " + Environment.NewLine
                                    + "            WHEN P.TipoIdentificador = 1 THEN 'PTI' " + Environment.NewLine
                                    + "            WHEN P.TipoIdentificador = 2 THEN 'PSS' " + Environment.NewLine
                                    + "            END AS TipoIdentificador, IdPSS, " + Environment.NewLine
                                    + "        CASE " + Environment.NewLine
                                    + "            WHEN P.EtapaProjetoId =  3 THEN [NomeProjeto] + ' - Refinamento' " + Environment.NewLine
                                    + "            ELSE [NomeProjeto] END AS 'NomeProjeto' " + Environment.NewLine
                                    + "        ,u.NomeUsuario AS 'Responsavel', case when area.NomeAreaGestora is not null then area.NomeAreaGestora else '' end  as GerenciaDemandante, CASE WHEN S.NomeUsuario IS NOT NULL THEN S.NomeUsuario  ELSE '' END as 'Solicitante',CONVERT(varchar(10), P.DataAtualizacao, 103) as 'DataAtualizacao',B.DataCadatrado " + Environment.NewLine
                                    + "        FROM [PORTAL_CORPORATIVO_PRD].[dbo].[Projetos] P " + Environment.NewLine
                                    + "        LEFT JOIN [dbo].[Usuarios] U on P.AnalistaResponsavelId = U.UsuarioId " + Environment.NewLine
                                    + "        LEFT JOIN  [dbo].[Usuarios] S on P.SolicitanteId = S.UsuarioId " + Environment.NewLine
                                    + "        LEFT JOIN[dbo].[AreasGestoras] area On S.AreaGestoraId = area.Id " + Environment.NewLine
                                    + "        LEFT JOIN(" + Environment.NewLine
                                    + "                  SELECT [ProjetoId], Max([Inicio]) DataCadatrado " + Environment.NewLine
                                    + "                  FROM [PORTAL_CORPORATIVO_PRD].[dbo].[ArquivosBiWeekly] A " + Environment.NewLine
                                    + "                  WHERE [TipoProjeto] = 2  " + Environment.NewLine
                                    + "                  GROUP BY  [ProjetoId]  " + Environment.NewLine
                                    + "                 ) B ON P.Id = B.ProjetoId " + Environment.NewLine
                                    + "          WHERE EtapaProjetoId IN (3,5,6,7,8,9) AND Deleted = 0 " + Environment.NewLine
                                    + "          GROUP BY P.[Id],[NomeProjeto], u.NomeUsuario,B.DataCadatrado, P.EtapaProjetoId,P.TipoIdentificador, IdPSS, u.NomeUsuario,B.DataCadatrado, p.DataAberturaPSS, p.DataCadastro, S.NomeUsuario, area.NomeAreaGestora, P.DataAtualizacao ) T " + Environment.NewLine
                                    + "LEFT JOIN [PORTAL_CORPORATIVO_PRD].[dbo].[ArquivosBiWeekly] Arquivo on Arquivo.ProjetoId = T.Id AND Arquivo.Inicio = T.DataCadatrado " + Environment.NewLine
                                    + "WHERE NomeProjeto NOT LIKE '%infra%' and NomeArquivo is not null";

                SqlCommand cmd = Util.sqlReader(sqlConsult);
                SqlDataReader rdr = cmd.ExecuteReader();

                if (rdr.HasRows)
                {
                    int line = 1;
                    var idPagina = ObterPosicaoDoSlideLista(caminhoFichaGerada, "12 Estruturantes EM ANDAMENTO");
                    
                    while (rdr.Read())
                    {
                        if (rdr["NomeArquivo"].ToString() != null)
                        {
                            byte[] imagemBytes = (byte[])rdr["Bytes"];
                            string nomeOriginal = rdr["Id"].ToString() + "_" + rdr["NomeArquivo"].ToString();
                            string extensao = System.IO.Path.GetExtension(nomeOriginal);

                            valoresFotos.Add(rdr["NomeProjeto"].ToString() + ";" + nomeOriginal);

                            // Define nome do arquivo salvo
                            string caminhoCompleto = System.IO.Path.Combine(folderPathImagemEstruturante, nomeOriginal);

                            // Salva a imagem
                            File.WriteAllBytes(caminhoCompleto, imagemBytes);
                        }

                        var valoresLinha = new List<string> { rdr["TipoIdentificador"].ToString() + " " + rdr["IdPSS"].ToString() + " - " + rdr["NomeProjeto"].ToString(), rdr["GerenciaDemandante"].ToString(), rdr["Solicitante"].ToString(), rdr["DataAtualizacao"].ToString() };
                        if (line >= 17)
                        {
                            
                            AdicionarLinha2(caminhoFichaGerada, valoresLinha, "12 Estruturantes EM ANDAMENTO", idPagina[1]);
                        }
                        else
                            AdicionarLinha2(caminhoFichaGerada, valoresLinha, "12 Estruturantes EM ANDAMENTO", idPagina[0]);
                        line++;
                    }
                    rdr.Close();
                    cmd.Connection.Close();
                    var idPaginaPrincipal = ObterPosicaoDoSlide(caminhoFichaGerada, "12 Estruturantes EM ANDAMENTO");

                    for (int i = 0; valoresFotos.Count > 0; i++)
                    {
                        var Photo = valoresFotos[i].Split(';');
                        AdicionarImagem.InserirSlideComImagemNaOrdem(caminhoFichaGerada, folderPathImagemEstruturante + @"\" + Photo[1], Photo[0], (idPaginaPrincipal + idPagina.Count) + (i));
                    }
                }
            }
            catch (Exception ex) { }
        }

        private void geraEstruturanteInfra()
        {
            try
            {
                List<string> valoresFotos = new List<string>();
                string sqlConsult = "SELECT T.Id, T.NomeProjeto, T.TipoIdentificador, T.IdPSS, T.GerenciaDemandante, T.Solicitante, T.DataAtualizacao, T.Responsavel, T.DataCadatrado, ContentType, Bytes, NomeArquivo " + Environment.NewLine
                                    + "FROM (" + Environment.NewLine
                                    + "        SELECT  P.[Id], " + Environment.NewLine
                                    + "        CASE " + Environment.NewLine
                                    + "            WHEN P.TipoIdentificador = 1 THEN 'PTI' " + Environment.NewLine
                                    + "            WHEN P.TipoIdentificador = 2 THEN 'PSS' " + Environment.NewLine
                                    + "            END AS TipoIdentificador, IdPSS, " + Environment.NewLine
                                    + "        CASE " + Environment.NewLine
                                    + "            WHEN P.EtapaProjetoId =  3 THEN [NomeProjeto] + ' - Refinamento' " + Environment.NewLine
                                    + "            ELSE [NomeProjeto] END AS 'NomeProjeto' " + Environment.NewLine
                                    + "        ,u.NomeUsuario AS 'Responsavel', case when area.NomeAreaGestora is not null then area.NomeAreaGestora else '' end  as GerenciaDemandante, CASE WHEN S.NomeUsuario IS NOT NULL THEN S.NomeUsuario  ELSE '' END as 'Solicitante',CONVERT(varchar(10), P.DataAtualizacao, 103) as 'DataAtualizacao',B.DataCadatrado " + Environment.NewLine
                                    + "        FROM [PORTAL_CORPORATIVO_PRD].[dbo].[Projetos] P " + Environment.NewLine
                                    + "        LEFT JOIN [dbo].[Usuarios] U on P.AnalistaResponsavelId = U.UsuarioId " + Environment.NewLine
                                    + "        LEFT JOIN  [dbo].[Usuarios] S on P.SolicitanteId = S.UsuarioId " + Environment.NewLine
                                    + "        LEFT JOIN[dbo].[AreasGestoras] area On S.AreaGestoraId = area.Id " + Environment.NewLine
                                    + "        LEFT JOIN(" + Environment.NewLine
                                    + "                  SELECT [ProjetoId], Max([Inicio]) DataCadatrado " + Environment.NewLine
                                    + "                  FROM [PORTAL_CORPORATIVO_PRD].[dbo].[ArquivosBiWeekly] A " + Environment.NewLine
                                    + "                  WHERE [TipoProjeto] = 2  " + Environment.NewLine
                                    + "                  GROUP BY  [ProjetoId]  " + Environment.NewLine
                                    + "                 ) B ON P.Id = B.ProjetoId " + Environment.NewLine
                                    + "          WHERE EtapaProjetoId IN (3,5,6,7,8,9) AND Deleted = 0 " + Environment.NewLine
                                    + "          GROUP BY P.[Id],[NomeProjeto], u.NomeUsuario,B.DataCadatrado, P.EtapaProjetoId,P.TipoIdentificador, IdPSS, u.NomeUsuario,B.DataCadatrado, p.DataAberturaPSS, p.DataCadastro, S.NomeUsuario, area.NomeAreaGestora, P.DataAtualizacao ) T " + Environment.NewLine
                                    + "LEFT JOIN [PORTAL_CORPORATIVO_PRD].[dbo].[ArquivosBiWeekly] Arquivo on Arquivo.ProjetoId = T.Id AND Arquivo.Inicio = T.DataCadatrado " + Environment.NewLine
                                    + "WHERE NomeProjeto LIKE '%infra%' ";

                SqlCommand cmd = Util.sqlReader(sqlConsult);
                SqlDataReader rdr = cmd.ExecuteReader();

                if (rdr.HasRows)
                {
                    int line = 1;
                    var idPagina = ObterPosicaoDoSlideLista(caminhoFichaGerada, "09 Melhorias de Infraestrutura com Estruturantes EM ANDAMENTO");

                    while (rdr.Read())
                    {
                        if (rdr["NomeArquivo"].ToString() != null)
                        {
                            byte[] imagemBytes = (byte[])rdr["Bytes"];
                            string nomeOriginal = rdr["Id"].ToString() + "_" + rdr["NomeArquivo"].ToString();
                            string extensao = System.IO.Path.GetExtension(nomeOriginal);

                            valoresFotos.Add(rdr["NomeProjeto"].ToString() + ";" + nomeOriginal);

                            // Define nome do arquivo salvo
                            string caminhoCompleto = System.IO.Path.Combine(folderPathImagemEstruturante, nomeOriginal);

                            // Salva a imagem
                            File.WriteAllBytes(caminhoCompleto, imagemBytes);
                        }

                        var valoresLinha = new List<string> { rdr["TipoIdentificador"].ToString() + " " + rdr["IdPSS"].ToString() + " - " + rdr["NomeProjeto"].ToString(), rdr["GerenciaDemandante"].ToString(), rdr["Solicitante"].ToString(), rdr["DataAtualizacao"].ToString() };
                        if (line >= 17)
                        {

                            AdicionarLinha2(caminhoFichaGerada, valoresLinha, "09 Melhorias de Infraestrutura com Estruturantes EM ANDAMENTO", idPagina[1]);
                        }
                        else
                            AdicionarLinha2(caminhoFichaGerada, valoresLinha, "09 Melhorias de Infraestrutura com Estruturantes EM ANDAMENTO", idPagina[0]);
                        line++;
                    }
                    rdr.Close();
                    cmd.Connection.Close();
                    var idPaginaPrincipal = ObterPosicaoDoSlide(caminhoFichaGerada, "09 Melhorias de Infraestrutura com Estruturantes EM ANDAMENTO");

                    for (int i = 0; valoresFotos.Count > 0; i++)
                    {
                        var Photo = valoresFotos[i].Split(';');
                        AdicionarImagem.InserirSlideComImagemNaOrdem(caminhoFichaGerada, folderPathImagemEstruturante + @"\" + Photo[1], Photo[0], (idPaginaPrincipal + idPagina.Count) + (i));
                    }
                }
            }
            catch (Exception ex) { }
        }
        private void geraProjetosExternos()
        {
            try
            {
                List<string> valoresFotos = new List<string>();
                string sqlConsult = "SELECT T.Id, NomeProjeto, GerenciaDemandante, Solicitante, DataAtualizacao, Responsavel, DataCadatrado, ContentType, Bytes, NomeArquivo " + Environment.NewLine 
                                    + " FROM ( " + Environment.NewLine
                                    + "     SELECT   P.[Id],NomeProjeto,  area.NomeAreaGestora as GerenciaDemandante, S.NomeUsuario as 'Solicitante',CONVERT(varchar(10), P.DataAtualizacao, 103) as 'DataAtualizacao', " + Environment.NewLine
                                    + "     u.NomeUsuario as 'Responsavel' , B.DataCadatrado  " + Environment.NewLine
                                    + "     FROM [PORTAL_CORPORATIVO_PRD].[dbo].[ProjetoExterno] P  " + Environment.NewLine
                                    + "     LEFT JOIN [dbo].[Usuarios] U on P.AnalistaResponsavelId = U.UsuarioId  " + Environment.NewLine
                                    + "     LEFT JOIN  [dbo].[Usuarios] S on P.SolicitanteId = S.UsuarioId  " + Environment.NewLine
                                    + "     LEFT JOIN [dbo].[AreasGestoras] area On S.AreaGestoraId = area.Id  " + Environment.NewLine
                                    + "     LEFT JOIN  " + Environment.NewLine
                                    + "     ( " + Environment.NewLine
                                    + "         SELECT [ProjetoId], Max([Inicio]) DataCadatrado  " + Environment.NewLine
                                    + "         FROM [PORTAL_CORPORATIVO_PRD].[dbo].[ArquivosBiWeekly] A  " + Environment.NewLine
                                    + "         WHERE [TipoProjeto] = 3 " + Environment.NewLine
                                    + "         GROUP BY  [ProjetoId]  " + Environment.NewLine
                                    + "     ) B on P.Id = B.ProjetoId  " + Environment.NewLine
                                    + "     WHERE   EtapaProjetoExternoId not in (8,20,31,47,63,75,76) AND Deleted = 0  and P.StatusProjetoExternoId = 3   " + Environment.NewLine
                                    + "     GROUP BY  P.[Id],NomeProjeto, u.NomeUsuario,B.DataCadatrado,   p.DataCadastro, S.NomeUsuario, area.NomeAreaGestora, P.DataAtualizacao  " + Environment.NewLine
                                    + " )T  " + Environment.NewLine
                                    + " LEFT JOIN [PORTAL_CORPORATIVO_PRD].[dbo].[ArquivosBiWeekly] Arquivo on Arquivo.ProjetoId = T.Id and Arquivo.Inicio = T.DataCadatrado  " + Environment.NewLine
                                    + " WHERE NomeArquivo is not null " + Environment.NewLine
                                    + " ORDER BY T.Id ASC";

                SqlCommand cmd = Util.sqlReader(sqlConsult);
                SqlDataReader rdr = cmd.ExecuteReader();

                if (rdr.HasRows)
                {
                    int line = 1;
                    var idPagina = ObterPosicaoDoSlideLista(caminhoFichaGerada, "Projetos de Negócio EM ANDAMENTO");

                    while (rdr.Read())
                    {
                        var sasa = rdr["NomeArquivo"].ToString();
                        if (rdr["NomeArquivo"].ToString() != "")
                        {
                            byte[] imagemBytes = (byte[])rdr["Bytes"];
                            string nomeOriginal = rdr["Id"].ToString() + "_" + rdr["NomeArquivo"].ToString();
                            string extensao = System.IO.Path.GetExtension(nomeOriginal);

                            valoresFotos.Add(rdr["NomeProjeto"].ToString() + ";" + nomeOriginal);

                            // Define nome do arquivo salvo
                            string caminhoCompleto = System.IO.Path.Combine(folderPathImagemEstruturante, nomeOriginal);

                            // Salva a imagem
                            File.WriteAllBytes(caminhoCompleto, imagemBytes);
                        }

                        var valoresLinha = new List<string> { rdr["NomeProjeto"].ToString(), rdr["GerenciaDemandante"].ToString(), rdr["Solicitante"].ToString(), rdr["DataAtualizacao"].ToString() };
                        if (line >= 17 && line <= 33)
                        {

                            AdicionarLinha2(caminhoFichaGerada, valoresLinha, "Projetos de Negócio EM ANDAMENTO", idPagina[1]);
                        }
                        else if (line >= 34 && line <= 50)
                        {

                            AdicionarLinha2(caminhoFichaGerada, valoresLinha, "Projetos de Negócio EM ANDAMENTO", idPagina[2]);
                        }
                        else if (line >= 51)
                        {

                            AdicionarLinha2(caminhoFichaGerada, valoresLinha, "Projetos de Negócio EM ANDAMENTO", idPagina[3]);
                        }

                        else
                            AdicionarLinha2(caminhoFichaGerada, valoresLinha, "Projetos de Negócio EM ANDAMENTO", idPagina[0]);
                        line++;
                    }
                    rdr.Close();
                    cmd.Connection.Close();
                    var idPaginaPrincipal = ObterPosicaoDoSlide(caminhoFichaGerada, "Projetos de Negócio EM ANDAMENTO");

                    for (int i = 0; valoresFotos.Count > 0; i++)
                    {
                        var Photo = valoresFotos[i].Split(';');
                        AdicionarImagem.InserirSlideComImagemNaOrdem(caminhoFichaGerada, folderPathImagemEstruturante + @"\" + Photo[1], Photo[0], (idPaginaPrincipal + idPagina.Count) + (i));
                    }
                }
            }
            catch (Exception ex) { }
        }


        private void geraProjetosEletronicos()
        {
            try
            {
                List<string> valoresFotos = new List<string>();
                string sqlConsult = "SELECT T.Id, NomeProjeto, GerenciaDemandante, Solicitante, DataAtualizacao, Responsavel, DataCadatrado, ContentType, Bytes, NomeArquivo  "+ Environment.NewLine
                                    + " FROM("+ Environment.NewLine
                                    + "     SELECT   P.[Id], NomeProjeto, area.NomeAreaGestora as GerenciaDemandante, S.NomeUsuario as 'Solicitante', CONVERT(varchar(10), P.DataAtualizacao, 103) as 'DataAtualizacao', " + Environment.NewLine
                                    + "     u.NomeUsuario as 'Responsavel', B.DataCadatrado  " + Environment.NewLine
                                    + "     FROM[PORTAL_CORPORATIVO_PRD].[dbo].[ProjetoEletronico] P  " + Environment.NewLine
                                    + "     LEFT JOIN[dbo].[Usuarios] U on P.AnalistaResponsavelId = U.UsuarioId  " + Environment.NewLine
                                    + "     LEFT JOIN[dbo].[Usuarios] S on P.SolicitanteId = S.UsuarioId  " + Environment.NewLine
                                    + "     LEFT JOIN[dbo].[AreasGestoras] area On S.AreaGestoraId = area.Id  " + Environment.NewLine
                                    + "     LEFT JOIN  " + Environment.NewLine
                                    + "     (" + Environment.NewLine
                                    + "         SELECT[ProjetoId], Max([Inicio]) DataCadatrado  " + Environment.NewLine
                                    + "         FROM[PORTAL_CORPORATIVO_PRD].[dbo].[ArquivosBiWeekly] A  " + Environment.NewLine
                                    + "         WHERE[TipoProjeto] = 4  " + Environment.NewLine
                                    + "         GROUP BY[ProjetoId]  " + Environment.NewLine
                                    + "     ) B on P.Id = B.ProjetoId  " + Environment.NewLine
                                    + "     WHERE   EtapaProjetoEletronicoId not in (6, 7, 13, 14, 20, 21, 27, 28, 34, 35) AND Deleted = 0  and P.StatusProjetoEletronicoId = 3  " + Environment.NewLine
                                    + "     GROUP BY  P.[Id], NomeProjeto, u.NomeUsuario, B.DataCadatrado, p.DataCadastro, S.NomeUsuario, area.NomeAreaGestora, P.DataAtualizacao  " + Environment.NewLine
                                    + " ) T  " + Environment.NewLine
                                    + " LEFT JOIN[PORTAL_CORPORATIVO_PRD].[dbo].[ArquivosBiWeekly] Arquivo on Arquivo.ProjetoId = T.Id and Arquivo.Inicio = T.DataCadatrado  " + Environment.NewLine
                                    + " WHERE NomeArquivo is not null  " + Environment.NewLine
                                    + " ORDER BY T.Id ASC";  

                SqlCommand cmd = Util.sqlReader(sqlConsult);
                SqlDataReader rdr = cmd.ExecuteReader();

                if (rdr.HasRows)
                {
                    int line = 1;
                    var idPagina = ObterPosicaoDoSlideLista(caminhoFichaGerada, "Projetos Eletrônicos EM ANDAMENTO");

                    while (rdr.Read())
                    {
                        var sasa = rdr["NomeArquivo"].ToString();
                        if (rdr["NomeArquivo"].ToString() != "")
                        {
                            byte[] imagemBytes = (byte[])rdr["Bytes"];
                            string nomeOriginal = rdr["Id"].ToString() + "_" + rdr["NomeArquivo"].ToString();
                            string extensao = System.IO.Path.GetExtension(nomeOriginal);

                            valoresFotos.Add(rdr["NomeProjeto"].ToString() + ";" + nomeOriginal);

                            // Define nome do arquivo salvo
                            string caminhoCompleto = System.IO.Path.Combine(folderPathImagemEletronico, nomeOriginal);

                            // Salva a imagem
                            File.WriteAllBytes(caminhoCompleto, imagemBytes);
                        }

                        var valoresLinha = new List<string> { rdr["NomeProjeto"].ToString(), rdr["GerenciaDemandante"].ToString(), rdr["Solicitante"].ToString(), rdr["DataAtualizacao"].ToString() };
                        if (line >= 17 && line <= 33)
                        {

                            AdicionarLinha2(caminhoFichaGerada, valoresLinha, "Projetos Eletrônicos EM ANDAMENTO", idPagina[1]);
                        }
                        else if (line >= 34 && line <= 50)
                        {

                            AdicionarLinha2(caminhoFichaGerada, valoresLinha, "Projetos Eletrônicos EM ANDAMENTO", idPagina[2]);
                        }
                        else if (line >= 51)
                        {

                            AdicionarLinha2(caminhoFichaGerada, valoresLinha, "Projetos Eletrônicos EM ANDAMENTO", idPagina[3]);
                        }

                        else
                            AdicionarLinha2(caminhoFichaGerada, valoresLinha, "Projetos Eletrônicos EM ANDAMENTO", idPagina[0]);
                        line++;
                    }
                    rdr.Close();
                    cmd.Connection.Close();
                    var idPaginaPrincipal = ObterPosicaoDoSlide(caminhoFichaGerada, "Projetos Eletrônicos EM ANDAMENTO");

                    for (int i = 0; valoresFotos.Count > 0; i++)
                    {
                        var Photo = valoresFotos[i].Split(';');
                        AdicionarImagem.InserirSlideComImagemNaOrdem(caminhoFichaGerada, folderPathImagemEletronico + @"\" + Photo[1], Photo[0], (idPaginaPrincipal + idPagina.Count) + (i));
                    }
                }
            }
            catch (Exception ex) { }
        }
        public static void AdicionarLinha2(string caminhoFicha, List<string> dadosTabela, string tabelaNome, string slideIndex)
        {
            try
            {
                using (PresentationDocument presentationDoc = PresentationDocument.Open(caminhoFicha, true))
                {
                    PresentationPart presentationPart = presentationDoc.PresentationPart;           

                    var slide = (SlidePart)presentationPart.GetPartById(slideIndex);

                    if (slide.Slide.InnerText.Contains(tabelaNome))
                    {
                        var tabela = slide.Slide.Descendants<Table>().FirstOrDefault();
                        if (tabela == null)
                        {
                            Console.WriteLine("Tabela não encontrada no slide especificado.");
                            return;
                        }
                        int numColunas = tabela.TableGrid.ChildElements.Count;
                        var novaLinha = new TableRow();
                        novaLinha.Height = 220000; // Altura da linha em EMUs
                                                   // Adiciona células à nova linha

                        foreach (var valor in dadosTabela)
                        {
                            var novaCelula = new TableCell();
                            // Adiciona o texto à célula
                            novaCelula.Append(
                                new DocumentFormat.OpenXml.Drawing.TextBody(
                                new BodyProperties(),
                                new ListStyle(),
                                //new Paragraph(new Run(new DocumentFormat.OpenXml.Drawing.Text(valor)))
                                new Paragraph(
                                              new Run(
                                                  new RunProperties() { FontSize = 10 * 100 }, // Define o tamanho da fonte
                                                  new DocumentFormat.OpenXml.Drawing.Text(valor)
                                                      )
                                              )
                            ));
                            // Define a margem padrão
                            novaCelula.Append(new TableCellProperties());

                            novaLinha.Append(novaCelula);
                        }
                        while (novaLinha.ChildElements.Count < numColunas)
                        {
                            novaLinha.Append(new TableCell(new TableCellProperties()));
                        }
                        // Adiciona a nova linha à tabela
                        tabela.Append(novaLinha);

                    }
                }
            }
            catch (Exception ex) { }
        }

        public static void AdicionarLinha(string caminhoFicha, List<string> dadosTabela, string tabelaNome, int slideIndex)
        {
            try
            {
                using (PresentationDocument presentationDoc = PresentationDocument.Open(caminhoFicha, true))
                {
                    PresentationPart presentationPart = presentationDoc.PresentationPart;
                    var slides = presentationDoc.PresentationPart.SlideParts;
                    foreach (var slide in slides)
                    {

                        if (slide.Slide.InnerText.Contains(tabelaNome))
                        {


                            // Obtém o slide desejado (começa em 0)
                            var slidePart = presentationDoc.PresentationPart.SlideParts.ElementAt(slideIndex);
                            // Busca a tabela no slide
                            var tabela = slide.Slide.Descendants<Table>().FirstOrDefault();
                            if (tabela == null)
                            {
                                Console.WriteLine("Tabela não encontrada no slide especificado.");
                                return;
                            }
                            // Obtém o número de colunas da tabela
                            int numColunas = tabela.TableGrid.ChildElements.Count;
                            // Cria uma nova linha
                            var novaLinha = new TableRow();
                            novaLinha.Height = 220000; // Altura da linha em EMUs
                                                       // Adiciona células à nova linha

                            foreach (var valor in dadosTabela)
                            {
                                var novaCelula = new TableCell();
                                // Adiciona o texto à célula
                                novaCelula.Append(
                                    new DocumentFormat.OpenXml.Drawing.TextBody(
                                    new BodyProperties(),
                                    new ListStyle(),
                                    //new Paragraph(new Run(new DocumentFormat.OpenXml.Drawing.Text(valor)))
                                    new Paragraph(
                                                  new Run(
                                                      new RunProperties() { FontSize = 10 * 100 }, // Define o tamanho da fonte
                                                      new DocumentFormat.OpenXml.Drawing.Text(valor)
                                                          )
                                                  )
                                ));
                                // Define a margem padrão
                                novaCelula.Append(new TableCellProperties());

                                novaLinha.Append(novaCelula);
                            }
                            // Garante que o número de células seja igual ao número de colunas
                            while (novaLinha.ChildElements.Count < numColunas)
                            {
                                novaLinha.Append(new TableCell(new TableCellProperties()));
                            }
                            // Adiciona a nova linha à tabela
                            tabela.Append(novaLinha);
                            break;
                        }
                    }
                }
            }
            catch (Exception ex) { }
        }

        public static int ObterPosicaoDoSlide(string caminhoPptx, string tituloSlide)
        {
            using (PresentationDocument presentationDoc = PresentationDocument.Open(caminhoPptx, false))
            {
                PresentationPart presentationPart = presentationDoc.PresentationPart;
                if (presentationPart == null || presentationPart.Presentation == null)
                    throw new InvalidDataException("O arquivo PPTX não possui uma estrutura válida.");
                var slideIdList = presentationPart.Presentation.SlideIdList;
                int posicao = 1; // A posição começa em 1
                foreach (var slideId in slideIdList.Elements<SlideId>())
                {
                    SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
                    // Verifica o título do slide (caso exista)
                    string titulo = ObterTituloDoSlide(slidePart);
                    if (!string.IsNullOrEmpty(titulo) && titulo.Contains(tituloSlide))
                    {
                        return posicao; // Retorna a posição do slide
                    }
                    posicao++;
                }
            }
            return -1; // Retorna -1 se o slide não for encontrado
        }

        private static string ObterTituloDoSlide(SlidePart slidePart)
        {
            try
            {
                if (slidePart == null || slidePart.Slide == null)
                    return null;
                var shape = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().FirstOrDefault(s => s.TextBody != null);
                if (shape.TextBody.InnerText != null)
                {
                    return shape.TextBody.InnerText.Trim();
                }
                else
                    return null;
            }
            catch { return null; }
        }

        public static List<string> ObterPosicaoDoSlideLista(string caminhoPptx, string tituloSlide)
        {
            List<string> ids = new List<string>();
            using (PresentationDocument presentationDoc = PresentationDocument.Open(caminhoPptx, true))
            {
                PresentationPart presentationPart = presentationDoc.PresentationPart;
                var ss = presentationPart.SlideParts.Count();

                var slideIdList = presentationPart.Presentation.SlideIdList;
                foreach (SlideId slideId in slideIdList)
                {
                    SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
                    if (slidePart.Slide.InnerText.Contains(tituloSlide))
                    {
                        ids.Add(slideId.RelationshipId);
                    }
                }
            }
            return ids;
        }

    }
}
