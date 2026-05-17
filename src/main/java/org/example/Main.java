package org.example;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Scanner;
import java.io.File;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
public class Main {
    public static void main(String[] args) {
        Scanner s = new Scanner(System.in); // Inicia a váriavel s como Scanner pra ser usada duarante o projeto.
        JFileChooser seletorArquivo = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter(
                "Selecione apenas arquivos do Excel.",
                "xlsx","xlsm"
        );
        seletorArquivo.setFileFilter(filter);
        File arquivoSelecionado = null;
        String r = "s";

        while (r.equals("s")) { // Loop pra sempre rodar o código abaixo.
            System.out.println("Digite o local do arquivo a ser desbloqueado:"); // Mensagem inicial para o usuario digitar o caminho do arquivo a desbloquear.
            int retorno = seletorArquivo.showOpenDialog(null);

            if (retorno == JFileChooser.APPROVE_OPTION){
                arquivoSelecionado = seletorArquivo.getSelectedFile();
            } else {
                System.out.println("Operação cancelada.");
                break;
            };

            // Pega apenas o nome do arquivo.
            // Exemplo: planilha.xlsx
            String nomeArquivoCompleto = arquivoSelecionado.getName();

            // Pega a pasta onde o arquivo está localizado.
            String localArquivo = arquivoSelecionado.getParent() + "\\";

            // Descobre a posição do último ponto no nome do arquivo.
            // Isso ajuda a separar nome e extensão.
            int indiceExtensao = nomeArquivoCompleto.lastIndexOf(".");

            // Pega apenas o nome do arquivo sem extensão.
            // Exemplo: planilha
            String nomeArquivo =
                    nomeArquivoCompleto.substring(0, indiceExtensao);

            // Pega apenas a extensão do arquivo.
            // Exemplo: .xlsx
            String extensaoArquivo =
                    nomeArquivoCompleto.substring(indiceExtensao);

            try { // Tenta executar o código abaixo

                // Cria um fluxo de entrada para ler o arquivo informado pelo usuário.
                // O FileInputStream é responsável por abrir o arquivo Excel na memória.
                FileInputStream arquivo = new FileInputStream(arquivoSelecionado);


                // Cria um objeto Workbook a partir do arquivo lido.
                // O XSSFWorkbook representa o arquivo Excel inteiro (.xlsx),
                // permitindo manipular abas, células, linhas, etc.
                XSSFWorkbook workbook = new XSSFWorkbook(arquivo);
                arquivo.close(); // Fecha a váriavel arquivo da memória

                int qtd_abas = workbook.getNumberOfSheets(); // Salva a quantidade de paginas na váriavel
                int index = 0; // Define a váriavel

                while (index < qtd_abas){

                    // Obtém a aba da planilha.
                    // O índice 0 representa a aba salva na váriavel.
                    XSSFSheet sheet = workbook.getSheetAt(index);

                    // Remove a proteção/bloqueio da aba selecionada.
                    sheet.disableLocking();

                    index++; // Adiciona um a váriavel Index

                }; // Loop pra desbloquear aba por aba.

                // Cria um fluxo de saída para salvar um novo arquivo Excel.
                // caminho.getParent() pega a pasta do arquivo original.
                // "\\Desbloqueado.xlsx" define o nome do novo arquivo.
                FileOutputStream novoArquivo =
                    new FileOutputStream(
                            localArquivo + nomeArquivo + "_desbloq" + extensaoArquivo
                    );

                workbook.write(novoArquivo); // Escreve o arquivo conforme o valor salvo no 'novoArquivo'

                novoArquivo.close();
                workbook.close();

                System.out.println("Arquivo Desbloqueado com Sucesso.");

            } catch (Exception e) { // Se der erro ele executa o código abaixo.
                System.out.println("Erro: "); // Escreve 'Erro: '
                e.printStackTrace(); // Especifica o Erro para o usuário.
            };

            System.out.println("\nDeseja continuar? [S/N]");
            r = s.nextLine().toLowerCase().substring(0);
        }

        System.out.println("Obrigado! Até mais.");
    }
}
