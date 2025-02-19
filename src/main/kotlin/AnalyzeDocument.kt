package org.example

import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.io.File
import java.io.FileInputStream

/**
 * @property filePath Wordファイルのパス
 */
class AnalyzeDocument(private var filePath: String) {
    private val document: XWPFDocument by lazy { loadDocument() }

    private fun loadDocument(): XWPFDocument {
        FileInputStream(File(filePath)).use { fis ->
            return XWPFDocument(fis)
        }
    }

    fun parseWordDocument(): DocumentContent {
        val paragraphs = document.paragraphs.map { it.text }
        val tables = document.tables.map { table ->
            table.rows.map { row ->
                row.tableCells.map { it.text }
            }
        }
        val images = document.allPictures.map { it.suggestFileExtension() }
        return DocumentContent(paragraphs, tables, images)
    }

    fun getFontSizes(): List<Double> {
        val fontSizes = mutableListOf<Double>()
        for (paragraph in document.paragraphs) {
            for (run in paragraph.runs) {
                val fontSize = run.fontSizeAsDouble
                if (fontSize > 0) {
                    fontSizes.add(fontSize)
                }
            }
        }
        return fontSizes
    }

    fun getPercent(): Long {
        return document.zoomPercent
    }
}