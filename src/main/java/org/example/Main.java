package org.example;

// Importações para manipulação de arquivos
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.File;
import java.util.Scanner;

// Biblioteca Apache POI para manipular Excel (.xlsx)
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.model.UnlockResult;
import org.example.service.ExcelService;

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
                "\n======================================================" +
                "\n            DESBLOQUEADOR DE PLANILHAS EXCEL         " +
                "\n======================================================" +
                "\n            Ferramenta de automação de arquivos       " +
                "\n======================================================"
        );

        // Loop principal do programa
        while (r.equals("s")) {

            System.out.println("\nAguardando seleção de arquivo...");

            // Abre janela para seleção do arquivo
            int retorno = seletorArquivo.showOpenDialog(null);

            // Verifica se o usuário selecionou um arquivo
            if (retorno == JFileChooser.APPROVE_OPTION){
                arquivoSelecionado = seletorArquivo.getSelectedFile();
                System.out.println("\nArquivo carregado com sucesso.");
                System.out.println("Arquivo: " + arquivoSelecionado.getName());
            } else {
                System.out.println("\nOperação cancelada pelo usuário.");
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

            UnlockResult resultado = ExcelService.unlock(arquivoSelecionado);

            if (resultado.isSuccess()) {
                System.out.println(resultado.getMessage());
            } else {
                System.out.println(resultado.getMessage());
            }



            // Pergunta ao usuário se deseja repetir o processo
            System.out.print("\nDeseja processar outro arquivo? [S/N]: ");

            // Lê resposta e normaliza para minúsculo
            String input = s.nextLine().toLowerCase().trim();
            r = input.isEmpty() ? "n" : input.substring(0, 1);
        }

        // Mensagem final do sistema
        System.out.print(
                "\n======================================================" +
                "\n                 PROCESSO FINALIZADO                 " +
                "\n              Obrigado por utilizar o sistema        " +
                "\n======================================================"
        );
    }
}