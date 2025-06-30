"use client"

import type React from "react"

import { useState, useRef } from "react"
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"
import { Badge } from "@/components/ui/badge"
import { Progress } from "@/components/ui/progress"
import { Upload, FileSpreadsheet, Calendar, Users, TrendingUp, Download, AlertCircle, CheckCircle } from "lucide-react"
import { Alert, AlertDescription } from "@/components/ui/alert"
import * as XLSX from "xlsx"

interface Employee {
  name: string
  adminStart?: Date
  adminEnd?: Date
  annualStart?: Date
  annualEnd?: Date
  originalIndex: number
}

interface WeeklyData {
  weekLabel: string
  workingEmployees: string[]
  weekStart: Date
  weekEnd: Date
}

interface ColumnMapping {
  name?: string
  adminStart?: string
  adminEnd?: string
  annualStart?: string
  annualEnd?: string
}

export default function ModernLeaveAnalyzer() {
  const [file, setFile] = useState<File | null>(null)
  const [employees, setEmployees] = useState<Employee[]>([])
  const [columns, setColumns] = useState<string[]>([])
  const [columnMapping, setColumnMapping] = useState<ColumnMapping>({ name: "" })
  const [startDate, setStartDate] = useState("2025-07-21")
  const [endDate, setEndDate] = useState("2025-09-08")
  const [weeklyData, setWeeklyData] = useState<WeeklyData[]>([])
  const [isAnalyzing, setIsAnalyzing] = useState(false)
  const [isGeneratingPDF, setIsGeneratingPDF] = useState(false)
  const [status, setStatus] = useState<"idle" | "loading" | "mapping" | "analyzed" | "error">("idle")
  const [errorMessage, setErrorMessage] = useState("")
  const fileInputRef = useRef<HTMLInputElement>(null)
  const reportRef = useRef<HTMLDivElement>(null)

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0]
    if (!selectedFile) return

    setFile(selectedFile)
    setStatus("loading")
    setErrorMessage("")

    try {
      const arrayBuffer = await selectedFile.arrayBuffer()
      const workbook = XLSX.read(arrayBuffer, { type: "array" })
      const sheetName = workbook.SheetNames[0]
      const worksheet = workbook.Sheets[sheetName]
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 })

      if (jsonData.length < 2) {
        throw new Error("Excel dosyası en az 2 satır içermelidir (başlık + veri)")
      }

      const headers = jsonData[0] as string[]
      const dataRows = jsonData.slice(1)

      setColumns(headers)

      // Auto-map columns with exact Turkish column names
      const autoMapping: ColumnMapping = { name: "" }
      headers.forEach((header) => {
        const headerTrimmed = header.trim()

        if (headerTrimmed === "İSİM") {
          autoMapping.name = header
        } else if (headerTrimmed === "İDARİ İZİN BAŞLAMA TARİHİ") {
          autoMapping.adminStart = header
        } else if (headerTrimmed === "İDARİ İZİN BİTİŞ TARİHİ") {
          autoMapping.adminEnd = header
        } else if (headerTrimmed === "YILLIK İZİN BAŞLAMA TARİHİ") {
          autoMapping.annualStart = header
        } else if (headerTrimmed === "YILLIK İZİN BİTİŞ TARİHİ") {
          autoMapping.annualEnd = header
        }
      })

      setColumnMapping(autoMapping)
      setStatus("mapping")
    } catch (error) {
      setErrorMessage(`Dosya yüklenirken hata: ${error instanceof Error ? error.message : "Bilinmeyen hata"}`)
      setStatus("error")
    }
  }

  const parseDate = (dateValue: any): Date | undefined => {
    if (!dateValue) return undefined

    if (dateValue instanceof Date) return dateValue

    if (typeof dateValue === "number") {
      // Excel date serial number - more accurate conversion
      const excelEpoch = new Date(1900, 0, 1)
      const date = new Date(excelEpoch.getTime() + (dateValue - 2) * 24 * 60 * 60 * 1000)
      return date
    }

    if (typeof dateValue === "string") {
      // Try different date formats
      const formats = [
        /^\d{1,2}[/\-.]\d{1,2}[/\-.]\d{4}$/, // DD/MM/YYYY or DD-MM-YYYY or DD.MM.YYYY
        /^\d{4}[/\-.]\d{1,2}[/\-.]\d{1,2}$/, // YYYY/MM/DD or YYYY-MM-DD or YYYY.MM.DD
      ]

      const dateStr = dateValue.toString().trim()

      // Try DD/MM/YYYY format first (Turkish format)
      if (formats[0].test(dateStr)) {
        const parts = dateStr.split(/[/\-.]/)
        const day = Number.parseInt(parts[0], 10)
        const month = Number.parseInt(parts[1], 10) - 1 // Month is 0-indexed
        const year = Number.parseInt(parts[2], 10)
        const date = new Date(year, month, day)
        if (!isNaN(date.getTime())) return date
      }

      // Try standard Date parsing
      const date = new Date(dateValue)
      if (!isNaN(date.getTime())) return date
    }

    return undefined
  }

  const processEmployeeData = async () => {
    if (!file || !columnMapping.name) {
      setErrorMessage("İsim sütunu seçilmesi zorunludur")
      return
    }

    try {
      const arrayBuffer = await file.arrayBuffer()
      const workbook = XLSX.read(arrayBuffer, { type: "array" })
      const sheetName = workbook.SheetNames[0]
      const worksheet = workbook.Sheets[sheetName]
      const jsonData = XLSX.utils.sheet_to_json(worksheet)

      const processedEmployees: Employee[] = []

      jsonData.forEach((row: any, index: number) => {
        const employee = {
          name: String(row[columnMapping.name!] || "").trim(),
          adminStart: columnMapping.adminStart ? parseDate(row[columnMapping.adminStart]) : undefined,
          adminEnd: columnMapping.adminEnd ? parseDate(row[columnMapping.adminEnd]) : undefined,
          annualStart: columnMapping.annualStart ? parseDate(row[columnMapping.annualStart]) : undefined,
          annualEnd: columnMapping.annualEnd ? parseDate(row[columnMapping.annualEnd]) : undefined,
          originalIndex: index, // Excel'deki orijinal sıra
        }

        // Debug log for first few employees
        if (processedEmployees.length < 3) {
          console.log("Processed employee:", {
            name: employee.name,
            adminStart: employee.adminStart?.toLocaleDateString("tr-TR"),
            adminEnd: employee.adminEnd?.toLocaleDateString("tr-TR"),
            annualStart: employee.annualStart?.toLocaleDateString("tr-TR"),
            annualEnd: employee.annualEnd?.toLocaleDateString("tr-TR"),
          })
        }

        if (employee.name) {
          processedEmployees.push(employee)
        }
      })

      setEmployees(processedEmployees)
      setStatus("analyzed")

      console.log(`Loaded ${processedEmployees.length} employees`)
    } catch (error) {
      setErrorMessage(`Veri işlenirken hata: ${error instanceof Error ? error.message : "Bilinmeyen hata"}`)
      setStatus("error")
    }
  }

  const isOnLeave = (employee: Employee, checkDate: Date): boolean => {
    // Normalize the check date to start of day for comparison
    const checkDateNormalized = new Date(checkDate.getFullYear(), checkDate.getMonth(), checkDate.getDate())

    // Check administrative leave
    if (employee.adminStart && employee.adminEnd) {
      const adminStartNormalized = new Date(
        employee.adminStart.getFullYear(),
        employee.adminStart.getMonth(),
        employee.adminStart.getDate(),
      )
      const adminEndNormalized = new Date(
        employee.adminEnd.getFullYear(),
        employee.adminEnd.getMonth(),
        employee.adminEnd.getDate(),
      )

      if (checkDateNormalized >= adminStartNormalized && checkDateNormalized <= adminEndNormalized) {
        return true
      }
    }

    // Check annual leave
    if (employee.annualStart && employee.annualEnd) {
      const annualStartNormalized = new Date(
        employee.annualStart.getFullYear(),
        employee.annualStart.getMonth(),
        employee.annualStart.getDate(),
      )
      const annualEndNormalized = new Date(
        employee.annualEnd.getFullYear(),
        employee.annualEnd.getMonth(),
        employee.annualEnd.getDate(),
      )

      if (checkDateNormalized >= annualStartNormalized && checkDateNormalized <= annualEndNormalized) {
        return true
      }
    }

    return false
  }

  const getWeekStart = (date: Date): Date => {
    const d = new Date(date)
    const day = d.getDay()
    // Monday = 1, Sunday = 0, so we want to get to Monday
    const diff = d.getDate() - day + (day === 0 ? -6 : 1)
    const monday = new Date(d.setDate(diff))
    // Reset time to start of day
    monday.setHours(0, 0, 0, 0)
    return monday
  }

  const analyzeData = () => {
    if (employees.length === 0) {
      setErrorMessage("Önce çalışan verilerini yükleyin")
      return
    }

    setIsAnalyzing(true)

    try {
      const start = new Date(startDate)
      const end = new Date(endDate)

      // Reset time to start of day
      start.setHours(0, 0, 0, 0)
      end.setHours(23, 59, 59, 999)

      const weeks: WeeklyData[] = []
      let currentWeekStart = getWeekStart(start)

      // Generate all weeks in the date range
      while (currentWeekStart <= end) {
        const weekEnd = new Date(currentWeekStart)
        weekEnd.setDate(weekEnd.getDate() + 6)
        weekEnd.setHours(23, 59, 59, 999)

        const workingEmployeesWithIndex: { name: string; originalIndex: number }[] = []

        employees.forEach((employee) => {
          let isWorkingThisWeek = false

          // Check each weekday (Monday to Friday)
          for (let dayOffset = 0; dayOffset < 5; dayOffset++) {
            const checkDate = new Date(currentWeekStart)
            checkDate.setDate(checkDate.getDate() + dayOffset)
            checkDate.setHours(12, 0, 0, 0) // Set to noon to avoid timezone issues

            // Only check dates within our analysis range
            if (checkDate >= start && checkDate <= end) {
              if (!isOnLeave(employee, checkDate)) {
                isWorkingThisWeek = true
                break // Employee is working at least one day this week
              }
            }
          }

          if (isWorkingThisWeek) {
            workingEmployeesWithIndex.push({
              name: employee.name,
              originalIndex: employee.originalIndex,
            })
          }
        })

        // Sort by original Excel order
        const workingEmployees = workingEmployeesWithIndex
          .sort((a, b) => a.originalIndex - b.originalIndex)
          .map((emp) => emp.name)

        // Format week label in Turkish
        const weekLabel =
          currentWeekStart.toLocaleDateString("tr-TR", {
            day: "numeric",
            month: "long",
            year: "numeric",
          }) + " Haftası"

        weeks.push({
          weekLabel,
          workingEmployees: workingEmployees, // .sort() kaldırıldı - Excel sırası korunuyor
          weekStart: new Date(currentWeekStart),
          weekEnd: new Date(weekEnd),
        })

        // Move to next week
        currentWeekStart = new Date(currentWeekStart)
        currentWeekStart.setDate(currentWeekStart.getDate() + 7)
      }

      setWeeklyData(weeks)
      setIsAnalyzing(false)

      // Log debug information
      console.log("Analysis completed:", {
        totalEmployees: employees.length,
        dateRange: `${start.toLocaleDateString("tr-TR")} - ${end.toLocaleDateString("tr-TR")}`,
        weeksAnalyzed: weeks.length,
        sampleEmployee: employees[0]
          ? {
              name: employees[0].name,
              adminStart: employees[0].adminStart?.toLocaleDateString("tr-TR"),
              adminEnd: employees[0].adminEnd?.toLocaleDateString("tr-TR"),
              annualStart: employees[0].annualStart?.toLocaleDateString("tr-TR"),
              annualEnd: employees[0].annualEnd?.toLocaleDateString("tr-TR"),
            }
          : null,
      })
    } catch (error) {
      setErrorMessage(`Analiz sırasında hata: ${error instanceof Error ? error.message : "Bilinmeyen hata"}`)
      setIsAnalyzing(false)
    }
  }

  const generateModernPDF = async () => {
    if (weeklyData.length === 0) {
      setErrorMessage("Önce analiz yapın")
      return
    }

    if (!reportRef.current) {
      setErrorMessage("Rapor elementi bulunamadı")
      return
    }

    setIsGeneratingPDF(true)

    try {
      // Import html2canvas dynamically
      const html2canvas = (await import("html2canvas")).default
      const jsPDF = (await import("jspdf")).default

      // Wait a bit for any animations to complete
      await new Promise((resolve) => setTimeout(resolve, 500))

      // Capture the report section with high quality
      const canvas = await html2canvas(reportRef.current, {
        scale: 2, // Higher quality
        useCORS: true,
        allowTaint: true,
        backgroundColor: "#ffffff",
        logging: false,
        width: reportRef.current.scrollWidth,
        height: reportRef.current.scrollHeight,
        windowWidth: 1200,
        windowHeight: reportRef.current.scrollHeight,
      })

      // Create PDF
      const pdf = new jsPDF("p", "mm", "a4")
      const imgWidth = 210 // A4 width in mm
      const imgHeight = (canvas.height * imgWidth) / canvas.width

      // If content is longer than one page, split it
      const pageHeight = 297 // A4 height in mm
      let heightLeft = imgHeight
      let position = 0

      // Add first page
      pdf.addImage(canvas.toDataURL("image/png"), "PNG", 0, position, imgWidth, imgHeight)
      heightLeft -= pageHeight

      // Add additional pages if needed
      while (heightLeft >= 0) {
        position = heightLeft - imgHeight
        pdf.addPage()
        pdf.addImage(canvas.toDataURL("image/png"), "PNG", 0, position, imgWidth, imgHeight)
        heightLeft -= pageHeight
      }

      // Save PDF
      const fileName = `Haftalik_Calisan_Raporu_${new Date().toISOString().split("T")[0]}.pdf`
      pdf.save(fileName)

      setIsGeneratingPDF(false)
    } catch (error) {
      console.error("PDF generation error:", error)
      setErrorMessage(`PDF oluşturulurken hata: ${error instanceof Error ? error.message : "Bilinmeyen hata"}`)
      setIsGeneratingPDF(false)
    }
  }

  const printReport = () => {
    if (weeklyData.length === 0) {
      setErrorMessage("Önce analiz yapın")
      return
    }

    // Open print dialog for the current page
    window.print()
  }

  const getStatusColor = (percentage: number) => {
    if (percentage >= 80) return "bg-green-500"
    if (percentage >= 60) return "bg-yellow-500"
    return "bg-red-500"
  }

  const getStatusText = (percentage: number) => {
    if (percentage >= 80) return "Yüksek"
    if (percentage >= 60) return "Orta"
    return "Düşük"
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 p-4">
      <div className="max-w-6xl mx-auto space-y-6">
        {/* Header */}
        <div className="text-center py-8 print:hidden">
          <h1 className="text-4xl font-bold text-gray-900 mb-2">Modern İzin Analiz Sistemi</h1>
          <p className="text-lg text-gray-600">Haftalık çalışan durumu analizi ve profesyonel raporlama</p>
        </div>

        {/* File Upload Card */}
        <Card className="shadow-lg print:hidden">
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              <FileSpreadsheet className="h-5 w-5" />
              Excel Dosyası Yükleme
            </CardTitle>
            <CardDescription>Çalışan izin verilerini içeren Excel dosyasını seçin</CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="flex items-center gap-4">
              <Input
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileUpload}
                ref={fileInputRef}
                className="hidden"
              />
              <Button
                onClick={() => fileInputRef.current?.click()}
                variant="outline"
                className="flex items-center gap-2"
              >
                <Upload className="h-4 w-4" />
                Dosya Seç
              </Button>
              {file && (
                <Badge variant="secondary" className="flex items-center gap-1">
                  <CheckCircle className="h-3 w-3" />
                  {file.name}
                </Badge>
              )}
            </div>

            {status === "error" && (
              <Alert variant="destructive">
                <AlertCircle className="h-4 w-4" />
                <AlertDescription>{errorMessage}</AlertDescription>
              </Alert>
            )}
          </CardContent>
        </Card>

        {/* Column Mapping Card */}
        {status === "mapping" && (
          <Card className="shadow-lg print:hidden">
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <Users className="h-5 w-5" />
                Sütun Eşleştirme
              </CardTitle>
              <CardDescription>Excel sütunlarını sistem alanlarıyla eşleştirin</CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-2">
                  <Label htmlFor="name-column">İsim Sütunu *</Label>
                  <Select
                    value={columnMapping.name || ""}
                    onValueChange={(value) => setColumnMapping((prev) => ({ ...prev, name: value }))}
                  >
                    <SelectTrigger>
                      <SelectValue placeholder="İsim sütununu seçin" />
                    </SelectTrigger>
                    <SelectContent>
                      {columns.map((col) => (
                        <SelectItem key={col} value={col}>
                          {col}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>

                <div className="space-y-2">
                  <Label htmlFor="admin-start">İdari İzin Başlangıç</Label>
                  <Select
                    value={columnMapping.adminStart ?? "__none__"}
                    onValueChange={(value) =>
                      setColumnMapping((prev) => ({
                        ...prev,
                        adminStart: value === "__none__" ? undefined : value,
                      }))
                    }
                  >
                    <SelectTrigger>
                      <SelectValue placeholder="Sütun seçin (opsiyonel)" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="__none__">Seçim yok</SelectItem>
                      {columns.map((col) => (
                        <SelectItem key={col} value={col}>
                          {col}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>

                <div className="space-y-2">
                  <Label htmlFor="admin-end">İdari İzin Bitiş</Label>
                  <Select
                    value={columnMapping.adminEnd ?? "__none__"}
                    onValueChange={(value) =>
                      setColumnMapping((prev) => ({
                        ...prev,
                        adminEnd: value === "__none__" ? undefined : value,
                      }))
                    }
                  >
                    <SelectTrigger>
                      <SelectValue placeholder="Sütun seçin (opsiyonel)" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="__none__">Seçim yok</SelectItem>
                      {columns.map((col) => (
                        <SelectItem key={col} value={col}>
                          {col}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>

                <div className="space-y-2">
                  <Label htmlFor="annual-start">Yıllık İzin Başlangıç</Label>
                  <Select
                    value={columnMapping.annualStart ?? "__none__"}
                    onValueChange={(value) =>
                      setColumnMapping((prev) => ({
                        ...prev,
                        annualStart: value === "__none__" ? undefined : value,
                      }))
                    }
                  >
                    <SelectTrigger>
                      <SelectValue placeholder="Sütun seçin (opsiyonel)" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="__none__">Seçim yok</SelectItem>
                      {columns.map((col) => (
                        <SelectItem key={col} value={col}>
                          {col}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>

                <div className="space-y-2">
                  <Label htmlFor="annual-end">Yıllık İzin Bitiş</Label>
                  <Select
                    value={columnMapping.annualEnd ?? "__none__"}
                    onValueChange={(value) =>
                      setColumnMapping((prev) => ({
                        ...prev,
                        annualEnd: value === "__none__" ? undefined : value,
                      }))
                    }
                  >
                    <SelectTrigger>
                      <SelectValue placeholder="Sütun seçin (opsiyonel)" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="__none__">Seçim yok</SelectItem>
                      {columns.map((col) => (
                        <SelectItem key={col} value={col}>
                          {col}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              </div>

              <Button onClick={processEmployeeData} disabled={!columnMapping.name} className="w-full">
                <CheckCircle className="h-4 w-4 mr-2" />
                Eşleştirmeyi Onayla
              </Button>
            </CardContent>
          </Card>
        )}

        {/* Analysis Settings Card */}
        {status === "analyzed" && (
          <Card className="shadow-lg print:hidden">
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <Calendar className="h-5 w-5" />
                Analiz Ayarları
              </CardTitle>
              <CardDescription>Analiz edilecek tarih aralığını belirleyin</CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-2">
                  <Label htmlFor="start-date">Başlangıç Tarihi</Label>
                  <Input id="start-date" type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} />
                </div>
                <div className="space-y-2">
                  <Label htmlFor="end-date">Bitiş Tarihi</Label>
                  <Input id="end-date" type="date" value={endDate} onChange={(e) => setEndDate(e.target.value)} />
                </div>
              </div>

              <div className="flex gap-4">
                <Button onClick={analyzeData} disabled={isAnalyzing} className="flex items-center gap-2">
                  <TrendingUp className="h-4 w-4" />
                  {isAnalyzing ? "Analiz Ediliyor..." : "Analiz Yap"}
                </Button>

                {weeklyData.length > 0 && (
                  <>
                    <Button
                      onClick={generateModernPDF}
                      disabled={isGeneratingPDF}
                      variant="outline"
                      className="flex items-center gap-2 bg-transparent"
                    >
                      <Download className="h-4 w-4" />
                      {isGeneratingPDF ? "PDF Oluşturuluyor..." : "PDF İndir"}
                    </Button>
                    <Button onClick={printReport} variant="outline" className="flex items-center gap-2 bg-transparent">
                      <Download className="h-4 w-4" />
                      Yazdır
                    </Button>
                  </>
                )}
              </div>

              <div className="text-sm text-gray-600">
                <p>Toplam {employees.length} çalışan yüklendi</p>
              </div>
            </CardContent>
          </Card>
        )}

        {/* Print Header - Only visible when printing */}
        {weeklyData.length > 0 && (
          <div className="hidden print:block text-center mb-8">
            <h1 className="text-2xl font-bold text-gray-900 mb-2">HAFTALİK ÇALIŞAN RAPORU</h1>
            <p className="text-sm text-gray-600">
              Analiz Dönemi: {new Date(startDate).toLocaleDateString("tr-TR")} -{" "}
              {new Date(endDate).toLocaleDateString("tr-TR")}
            </p>
            <p className="text-xs text-gray-500">
              Rapor Tarihi:{" "}
              {new Date().toLocaleDateString("tr-TR", {
                year: "numeric",
                month: "long",
                day: "numeric",
                hour: "2-digit",
                minute: "2-digit",
              })}
            </p>
          </div>
        )}

        {/* Results */}
        {weeklyData.length > 0 && (
          <div ref={reportRef} className="space-y-6">
            {/* Summary Statistics */}
            <Card className="shadow-lg print:shadow-none print:border">
              <CardHeader>
                <CardTitle>Özet İstatistikler</CardTitle>
              </CardHeader>
              <CardContent>
                <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                  <div className="text-center p-4 bg-blue-50 rounded-lg print:bg-blue-100">
                    <div className="text-2xl font-bold text-blue-600">{employees.length}</div>
                    <div className="text-sm text-gray-600">Toplam Çalışan</div>
                  </div>
                  <div className="text-center p-4 bg-green-50 rounded-lg print:bg-green-100">
                    <div className="text-2xl font-bold text-green-600">{weeklyData.length}</div>
                    <div className="text-sm text-gray-600">Analiz Edilen Hafta</div>
                  </div>
                  <div className="text-center p-4 bg-orange-50 rounded-lg print:bg-orange-100">
                    <div className="text-2xl font-bold text-orange-600">
                      {(
                        weeklyData.reduce((sum, week) => sum + week.workingEmployees.length, 0) / weeklyData.length
                      ).toFixed(1)}
                    </div>
                    <div className="text-sm text-gray-600">Ortalama Çalışan</div>
                  </div>
                  <div className="text-center p-4 bg-purple-50 rounded-lg print:bg-purple-100">
                    <div className="text-2xl font-bold text-purple-600">
                      %
                      {(
                        (weeklyData.reduce((sum, week) => sum + week.workingEmployees.length, 0) /
                          weeklyData.length /
                          employees.length) *
                        100
                      ).toFixed(1)}
                    </div>
                    <div className="text-sm text-gray-600">Ortalama Yoğunluk</div>
                  </div>
                </div>
              </CardContent>
            </Card>

            {/* Weekly Details */}
            <div className="grid gap-6">
              {weeklyData.map((week, index) => {
                const percentage = employees.length > 0 ? (week.workingEmployees.length / employees.length) * 100 : 0
                return (
                  <Card key={index} className="shadow-lg print:shadow-none print:border print:break-inside-avoid">
                    <CardHeader>
                      <div className="flex items-center justify-between">
                        <CardTitle className="text-lg">{week.weekLabel}</CardTitle>
                        <Badge variant="secondary" className={`${getStatusColor(percentage)} text-white print:border`}>
                          {getStatusText(percentage)}
                        </Badge>
                      </div>
                      <div className="space-y-2">
                        <div className="flex items-center justify-between text-sm">
                          <span>
                            Çalışan Sayısı: {week.workingEmployees.length}/{employees.length}
                          </span>
                          <span>%{percentage.toFixed(1)}</span>
                        </div>
                        <Progress value={percentage} className="h-2" />
                      </div>
                    </CardHeader>
                    <CardContent>
                      {week.workingEmployees.length > 0 ? (
                        <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-2 print:grid-cols-4">
                          {week.workingEmployees.map((employee, empIndex) => (
                            <div key={empIndex} className="text-sm p-2 bg-gray-50 rounded print:bg-gray-100">
                              {empIndex + 1}. {employee}
                            </div>
                          ))}
                        </div>
                      ) : (
                        <div className="text-center py-8 text-gray-500">
                          Bu hafta hiçbir çalışan aktif görevde bulunmamaktadır
                        </div>
                      )}
                    </CardContent>
                  </Card>
                )
              })}
            </div>
          </div>
        )}
      </div>

      <style jsx global>{`
        @media print {
          body {
            background: white !important;
          }
          .print\\:hidden {
            display: none !important;
          }
          .print\\:block {
            display: block !important;
          }
          .print\\:shadow-none {
            box-shadow: none !important;
          }
          .print\\:border {
            border: 1px solid #e5e7eb !important;
          }
          .print\\:break-inside-avoid {
            break-inside: avoid !important;
          }
          .print\\:bg-blue-100 {
            background-color: #dbeafe !important;
          }
          .print\\:bg-green-100 {
            background-color: #dcfce7 !important;
          }
          .print\\:bg-orange-100 {
            background-color: #fed7aa !important;
          }
          .print\\:bg-purple-100 {
            background-color: #f3e8ff !important;
          }
          .print\\:bg-gray-100 {
            background-color: #f3f4f6 !important;
          }
          .print\\:grid-cols-4 {
            grid-template-columns: repeat(4, minmax(0, 1fr)) !important;
          }
        }
      `}</style>
    </div>
  )
}
