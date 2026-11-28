package edu.sisinf.chuspadb.servicio;

import edu.sisinf.chuspadb.core.ErrorPersistencia;
import edu.sisinf.chuspadb.modelo.ResultadoQuery;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Objects;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * Servicio para exportar resultados de consultas a archivos Excel (.xlsx).
 */
public class ExportadorExcel {

    private static final Logger LOGGER = Logger.getLogger(ExportadorExcel.class.getName());

    /** Límite de filas de una hoja .xlsx (incluye la fila de encabezado). */
    private static final int MAX_FILAS_XLSX = 1_048_576;

    /**
     * Exporta un resultado de consulta a un archivo Excel.
     *
     * @param resultado resultado de la consulta (no nulo, de tipo LECTURA)
     * @param archivo   archivo destino (no nulo, no puede ser un directorio)
     * @throws NullPointerException     si algún argumento es nulo
     * @throws IllegalArgumentException si el archivo destino no es válido
     * @throws ErrorPersistencia        si el resultado es inconsistente o falla la escritura
     */
    public void exportar(ResultadoQuery resultado, File archivo) {
        Objects.requireNonNull(resultado, "El resultado a exportar no puede ser nulo");
        Objects.requireNonNull(archivo, "El archivo destino no puede ser nulo");

        validarArchivoDestino(archivo);

        if (resultado.getTipo() != ResultadoQuery.Tipo.LECTURA) {
            throw new ErrorPersistencia("Solo se pueden exportar resultados de lectura");
        }

        List<String> columnas = resultado.getColumnas();
        List<List<Object>> filas = resultado.getFilas();
        validarEstructura(columnas, filas);

        LOGGER.fine(() -> "Exportando " + filas.size() + " filas y " + columnas.size()
                + " columnas a " + archivo.getAbsolutePath());

        try (XSSFWorkbook workbook = new XSSFWorkbook();
             FileOutputStream out = new FileOutputStream(archivo)) {

            Sheet hoja = workbook.createSheet("Resultado");

            Row filaEncabezado = hoja.createRow(0);
            for (int c = 0; c < columnas.size(); c++) {
                String nombre = columnas.get(c);
                filaEncabezado.createCell(c).setCellValue(nombre == null ? "" : nombre);
            }

            for (int f = 0; f < filas.size(); f++) {
                Row fila = hoja.createRow(f + 1);
                List<Object> datosFila = filas.get(f);

                for (int c = 0; c < datosFila.size(); c++) {
                    Cell celda = fila.createCell(c);
                    escribirValor(celda, datosFila.get(c));
                }
            }

            for (int c = 0; c < columnas.size(); c++) {
                hoja.autoSizeColumn(c);
            }

            workbook.write(out);
            LOGGER.fine(() -> "Exportación a Excel completada: " + archivo.getAbsolutePath());

        } catch (IOException e) {
            LOGGER.log(Level.SEVERE, "Error al exportar archivo Excel: " + archivo.getAbsolutePath(), e);
            throw new ErrorPersistencia("Error al exportar archivo Excel", e);
        }
    }

    private void validarArchivoDestino(File archivo) {
        if (archivo.getName().isBlank()) {
            throw new IllegalArgumentException("El nombre del archivo destino no puede estar vacío");
        }
        if (archivo.isDirectory()) {
            throw new IllegalArgumentException(
                    "El destino es un directorio, no un archivo: " + archivo.getAbsolutePath());
        }
        File padre = archivo.getAbsoluteFile().getParentFile();
        if (padre != null && !padre.exists()) {
            throw new IllegalArgumentException(
                    "La carpeta destino no existe: " + padre.getAbsolutePath());
        }
        if (archivo.exists() && !archivo.canWrite()) {
            throw new IllegalArgumentException(
                    "No se puede escribir en el archivo destino: " + archivo.getAbsolutePath());
        }
    }

    private void validarEstructura(List<String> columnas, List<List<Object>> filas) {
        if (columnas == null) {
            throw new ErrorPersistencia("El resultado no contiene columnas (lista nula)");
        }
        if (filas == null) {
            throw new ErrorPersistencia("El resultado no contiene filas (lista nula)");
        }
        if (filas.size() + 1 > MAX_FILAS_XLSX) {
            throw new ErrorPersistencia("El resultado supera el máximo de filas de Excel ("
                    + (MAX_FILAS_XLSX - 1) + ")");
        }
        for (int f = 0; f < filas.size(); f++) {
            List<Object> fila = filas.get(f);
            if (fila == null) {
                throw new ErrorPersistencia("La fila " + (f + 1) + " es nula");
            }
            if (fila.size() != columnas.size()) {
                throw new ErrorPersistencia("La fila " + (f + 1) + " tiene " + fila.size()
                        + " valores, pero hay " + columnas.size() + " columnas");
            }
        }
    }

    private void escribirValor(Cell celda, Object valor) {
        Objects.requireNonNull(celda, "La celda no puede ser nula");

        if (valor == null) {
            celda.setBlank();
            return;
        }
        if (valor instanceof Number numero) {
            celda.setCellValue(numero.doubleValue());
            return;
        }
        if (valor instanceof Boolean booleano) {
            celda.setCellValue(booleano);
            return;
        }
        celda.setCellValue(valor.toString());
    }
}
