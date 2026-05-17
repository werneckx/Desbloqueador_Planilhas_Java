package org.example;

// Importações para manipulação de arquivos
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.File;
import java.util.Scanner;

// Biblioteca Apache POI para manipular Excel (.xlsx)
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// Interface gráfica para seleção de arquivos
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

public class Main {
public static void main(String[] args) {

        // Scanner para entrada do usuário via console
        Scanner s = new Scanner(System.in);

        // Componente gráfico para selecionar arquivos
        JFileChooser seletorArquivo = new JFileChooser();

        // Filtro para permitir apenas arquivos do Excel
        FileNameExtensionFilter filter = new FileNameExtensionFilter(
                "Arquivos do Excel",
                "xlsx", "xls", "xlsm", "xlsb", "xltx", "xltm", "xlt", "csv"
        );
        seletorArquivo.setFileFilter(filter);

        // Arquivo selecionado pelo usuário
        File arquivoSelecionado = null;

        // Controle do loop principal (continuar ou não)
        String r = "s";

        // Mensagem inicial do sistema
        System.out.print(
                "\n-----------------------------------------------" +
                "\n--- Bem vindo ao desbloqueador de planilhas ---" +
                "\n-----------------------------------------------"
        );

        // Loop principal do programa
        while (r.equals("s")) {

            // Abre janela para seleção do arquivo
            int retorno = seletorArquivo.showOpenDialog(null);

            // Verifica se o usuário selecionou um arquivo
            if (retorno == JFileChooser.APPROVE_OPTION){
                arquivoSelecionado = seletorArquivo.getSelectedFile();
                System.out.print("\nArquivo selecionado.");
            } else {
                System.out.println("\nOperação cancelada.");
                break;
            }

            // Obtém o nome completo do arquivo (com extensão)
            String nomeArquivoCompleto = arquivoSelecionado.getName();

            // Obtém o diretório onde o arquivo está localizado
            String localArquivo = arquivoSelecionado.getParent() + "\\";

            // Localiza o ponto da extensão
            int indiceExtensao = nomeArquivoCompleto.lastIndexOf(".");

            // Separa nome do arquivo (sem extensão)
            String nomeArquivo =
                    nomeArquivoCompleto.substring(0, indiceExtensao);

            // Separa apenas a extensão do arquivo
            String extensaoArquivo =
                    nomeArquivoCompleto.substring(indiceExtensao);

            try {
                // Abre o arquivo Excel para leitura
                FileInputStream arquivo = new FileInputStream(arquivoSelecionado);

                // Cria o workbook (arquivo Excel completo em memória)
                XSSFWorkbook workbook = new XSSFWorkbook(arquivo);

                // Fecha o fluxo de entrada após carregar o arquivo
                arquivo.close();

                // Percorre todas as abas da planilha
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {

                    // Obtém a aba atual
                    XSSFSheet sheet = workbook.getSheetAt(i);

                    // Remove proteção (desbloqueia a planilha)
                    sheet.disableLocking();
                }

                // Cria o arquivo de saída (novo arquivo desbloqueado)
                FileOutputStream novoArquivo =
                        new FileOutputStream(
                                localArquivo + nomeArquivo + "_desbloq" + extensaoArquivo
                        );

                // Escreve o conteúdo modificado no novo arquivo
                workbook.write(novoArquivo);

                // Fecha o fluxo de saída
                novoArquivo.close();

                // Fecha o workbook (libera memória)
                workbook.close();

                System.out.print("\nArquivo Desbloqueado com Sucesso.");

            } catch (Exception e) {
                // Captura qualquer erro durante o processo
                System.out.println("Erro: ");
                e.printStackTrace();
            }

            // Pergunta ao usuário se deseja repetir o processo
            System.out.print("\nDeseja continuar? [S/N]");

            // Lê resposta e normaliza para minúsculo
            r = s.nextLine().toLowerCase().length() > 1 ? r.substring(0,1): "n";
        }

        // Mensagem final do sistema
        System.out.print(
                "\n---------------------------" +
                "\n--- Obrigado! Até mais. ---" +
                "\n---------------------------"
        );
    }
}