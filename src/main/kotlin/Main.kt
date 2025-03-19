package org.example

import kotlinx.serialization.Serializable
import kotlinx.serialization.json.Json
import java.io.File

@Serializable
data class DocumentContent(
    val paragraphs: List<String>,
    val tables: List<List<List<String>>>,
    val images: List<String>
)

fun main() {
    val filePath = "src/main/resources/sample.docx"
    val analyzeDocument = AnalyzeDocument(filePath)
    val documentContent = analyzeDocument.parseWordDocument()
//    val fontSizes = analyzeDocument.getFontSizes()
    val percent = analyzeDocument.getPercent()
    val json = Json.encodeToString(documentContent)
    println(json)
//    println("Font sizes: $fontSizes")
    println("Zoom percent: $percent")
    File("src/main/resources/sample.json").writeText(json)
}
