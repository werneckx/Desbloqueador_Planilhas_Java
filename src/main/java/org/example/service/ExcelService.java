package org.example.service;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.model.UnlockResult;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import java.io.File;

public class ExcelService {
    public static UnlockResult unlock(
            File arquivoSelecionado
    ){
         try {

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

            // Abre o arquivo Excel para leitura
            FileInputStream arquivo = new FileInputStream(arquivoSelecionado);

            // Cria o workbook (arquivo Excel completo em memória)
            XSSFWorkbook workbook = new XSSFWorkbook(arquivo);

            // Percorre todas as abas da planilha
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {

                // Obtém a aba atual
                XSSFSheet sheet = workbook.getSheetAt(i);

                // Remove proteção (desbloqueia a planilha)
                sheet.disableLocking();
            }

            File arquivoGerado = new File(localArquivo + nomeArquivo + "_desbloqueado" + extensaoArquivo);

            // Cria o arquivo de saída (novo arquivo desbloqueado)
            FileOutputStream novoArquivo =
                    new FileOutputStream(
                            arquivoGerado
                    );

            // Escreve o conteúdo modificado no novo arquivo
            workbook.write(novoArquivo);

            UnlockResult resultado = new UnlockResult(
                    true,
                    "Arquivo desbloqueado com sucesso.",
                    arquivoGerado,
                    null
            );

            return resultado;

        } catch (Exception e) {
             UnlockResult resultado = new UnlockResult(
                     false,
                     "Erro!" + e.getMessage(),
                     null,
                     e.getMessage()
             );
            return resultado;
        }
    }
}
