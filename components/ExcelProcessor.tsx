'use client'

import { useState, useCallback, useEffect } from 'react'
import { useDropzone } from 'react-dropzone'
import * as XLSX from 'xlsx'

type SheetData = {
  [sheet: string]: string[]
}

export default function ExcelProcessor() {
  const [sourceFile, setSourceFile] = useState<File | null>(null)
  const [targetFile, setTargetFile] = useState<File | null>(null)
  const [sourceSheets, setSourceSheets] = useState<SheetData>({})
  const [targetSheets, setTargetSheets] = useState<SheetData>({})
  const [selectedSourceSheet, setSelectedSourceSheet] = useState<string>('')
  const [selectedTargetSheet, setSelectedTargetSheet] = useState<string>('')
  const [selectedSourceColumn, setSelectedSourceColumn] = useState<string>('')
  const [selectedTargetColumn, setSelectedTargetColumn] = useState<string>('')
  const [resultFile, setResultFile] = useState<string | null>(null)
  const [processing, setProcessing] = useState<boolean>(false)
  const [error, setError] = useState<string | null>(null)
  const [debug, setDebug] = useState<string | null>(null)

  // 重置所有状态的函数
  const resetAll = () => {
    setSourceFile(null)
    setTargetFile(null)
    setSourceSheets({})
    setTargetSheets({})
    setSelectedSourceSheet('')
    setSelectedTargetSheet('')
    setSelectedSourceColumn('')
    setSelectedTargetColumn('')
    setResultFile(null)
    setProcessing(false)
    setError(null)
    setDebug(null)
  }

  // 当源文件改变时，重置相关状态
  const onSourceDrop = useCallback((acceptedFiles: File[]) => {
    if (acceptedFiles && acceptedFiles.length > 0) {
      const file = acceptedFiles[0]
      setSourceFile(file)
      setSelectedSourceSheet('')
      setSelectedSourceColumn('')
      setResultFile(null)
      setError(null)
      readExcelFile(file, setSourceSheets)
    }
  }, [])

  // 当目标文件改变时，重置相关状态
  const onTargetDrop = useCallback((acceptedFiles: File[]) => {
    if (acceptedFiles && acceptedFiles.length > 0) {
      const file = acceptedFiles[0]
      setTargetFile(file)
      setSelectedTargetSheet('')
      setSelectedTargetColumn('')
      setResultFile(null)
      setError(null)
      readExcelFile(file, setTargetSheets)
    }
  }, [])

  // 当源工作表改变时，重置源列选择
  const handleSourceSheetChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const sheetName = e.target.value
    setSelectedSourceSheet(sheetName)
    setSelectedSourceColumn('')
    setResultFile(null)
  }

  // 当目标工作表改变时，重置目标列选择
  const handleTargetSheetChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const sheetName = e.target.value
    setSelectedTargetSheet(sheetName)
    setSelectedTargetColumn('')
    setResultFile(null)
  }

  const {
    getRootProps: getSourceRootProps,
    getInputProps: getSourceInputProps,
  } = useDropzone({
    onDrop: onSourceDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls'],
    },
    maxFiles: 1,
  })

  const {
    getRootProps: getTargetRootProps,
    getInputProps: getTargetInputProps,
  } = useDropzone({
    onDrop: onTargetDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls'],
    },
    maxFiles: 1,
  })

  const readExcelFile = (file: File, setSheetData: React.Dispatch<React.SetStateAction<SheetData>>) => {
    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer)
        const workbook = XLSX.read(data, { type: 'array', cellDates: true })
        
        const sheetsData: SheetData = {}
        
        workbook.SheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName]
          const headers: string[] = []
          
          // 获取列标题（假设在第一行）
          const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1')
          for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col })
            const cell = worksheet[cellAddress]
            const colName = cell ? String(cell.v) : XLSX.utils.encode_col(col)
            headers.push(colName)
          }
          
          sheetsData[sheetName] = headers
        })
        
        setSheetData(sheetsData)
      } catch (err) {
        setError('读取Excel文件时出错: ' + (err instanceof Error ? err.message : String(err)))
        console.error(err)
      }
    }
    reader.readAsArrayBuffer(file)
  }

  const processExcelFiles = () => {
    if (!sourceFile || !targetFile || !selectedSourceSheet || !selectedTargetSheet || 
        !selectedSourceColumn || !selectedTargetColumn) {
      setError('请选择所有必要的选项')
      return
    }

    setProcessing(true)
    setError(null)
    setDebug(null)

    try {
      const sourceReader = new FileReader()
      sourceReader.onload = (e) => {
        try {
          const sourceData = new Uint8Array(e.target?.result as ArrayBuffer)
          const sourceWorkbook = XLSX.read(sourceData, { type: 'array', cellDates: true })
          const sourceWorksheet = sourceWorkbook.Sheets[selectedSourceSheet]
          
          const targetReader = new FileReader()
          targetReader.onload = (e2) => {
            try {
              const targetData = new Uint8Array(e2.target?.result as ArrayBuffer)
              const targetWorkbook = XLSX.read(targetData, { type: 'array', cellDates: true })
              const targetWorksheet = targetWorkbook.Sheets[selectedTargetSheet]
              
              // 获取源列的索引
              const sourceColumnIndex = sourceSheets[selectedSourceSheet].indexOf(selectedSourceColumn)
              // 获取目标列的索引
              const targetColumnIndex = targetSheets[selectedTargetSheet].indexOf(selectedTargetColumn)
              
              if (sourceColumnIndex === -1 || targetColumnIndex === -1) {
                setError('无法找到选定的列，请重新选择')
                setProcessing(false)
                return
              }
              
              // 解析源工作表数据
              const sourceJsonData = XLSX.utils.sheet_to_json<any[]>(sourceWorksheet, { header: 1, defval: null })
              // 解析目标工作表数据
              const targetJsonData = XLSX.utils.sheet_to_json<any[]>(targetWorksheet, { header: 1, defval: null })
              
              // 记录调试信息
              const debugInfo = {
                sourceColumnIndex,
                targetColumnIndex,
                sourceRowCount: sourceJsonData.length,
                targetRowCount: targetJsonData.length,
              }
              
              // 复制源列数据到目标列
              let copiedCount = 0;
              for (let i = 1; i < sourceJsonData.length; i++) {
                const sourceRow = sourceJsonData[i]
                
                if (i >= targetJsonData.length) {
                  // 如果目标行不存在，创建新行
                  const newRow: any[] = []
                  targetJsonData.push(newRow)
                }
                
                if (sourceRow && sourceColumnIndex < sourceRow.length) {
                  const sourceValue = sourceRow[sourceColumnIndex]
                  
                  // 确保目标行存在
                  if (!Array.isArray(targetJsonData[i])) {
                    targetJsonData[i] = []
                  }
                  
                  // 复制值到目标列
                  targetJsonData[i][targetColumnIndex] = sourceValue
                  copiedCount++
                }
              }
              
              // 将修改后的数据转回工作表
              const newWorksheet = XLSX.utils.aoa_to_sheet(targetJsonData)
              
              // 保留原始工作表的样式和合并单元格信息
              if (targetWorksheet['!merges']) {
                newWorksheet['!merges'] = targetWorksheet['!merges']
              }
              if (targetWorksheet['!cols']) {
                newWorksheet['!cols'] = targetWorksheet['!cols']
              }
              if (targetWorksheet['!rows']) {
                newWorksheet['!rows'] = targetWorksheet['!rows']
              }
              
              targetWorkbook.Sheets[selectedTargetSheet] = newWorksheet
              
              // 生成新的Excel文件
              const outputData = XLSX.write(targetWorkbook, { 
                bookType: 'xlsx', 
                type: 'array',
                cellDates: true
              })
              
              // 创建Blob并生成下载链接
              const blob = new Blob([outputData], { type: 'application/octet-stream' })
              const url = URL.createObjectURL(blob)
              setResultFile(url)
              setDebug(`成功复制了 ${copiedCount} 行数据，从列 ${selectedSourceColumn} 到列 ${selectedTargetColumn}`)
              setProcessing(false)
            } catch (err) {
              setError('处理目标Excel文件时出错: ' + (err instanceof Error ? err.message : String(err)))
              setProcessing(false)
              console.error(err)
            }
          }
          targetReader.readAsArrayBuffer(targetFile)
        } catch (err) {
          setError('处理源Excel文件时出错: ' + (err instanceof Error ? err.message : String(err)))
          setProcessing(false)
          console.error(err)
        }
      }
      sourceReader.readAsArrayBuffer(sourceFile)
    } catch (err) {
      setError('处理文件时出错: ' + (err instanceof Error ? err.message : String(err)))
      setProcessing(false)
      console.error(err)
    }
  }

  return (
    <div className="max-w-5xl mx-auto bg-white px-6 py-10">
      <div className="flex justify-between items-center mb-10">
        <h1 className="text-2xl font-normal text-gray-800">Excel列复制工具</h1>
        <button 
          onClick={resetAll}
          className="flex items-center justify-center w-10 h-10 rounded-full hover:bg-gray-100 transition-colors"
          title="重置"
        >
          <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 text-gray-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
          </svg>
        </button>
      </div>

      <p className="text-center text-gray-500 mb-12 max-w-2xl mx-auto">选择源Excel文件和目标Excel文件，然后指定要复制的列</p>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-10 mb-12">
        {/* 源文件上传区域 */}
        <div className="space-y-6">
          <div className="flex items-center mb-4">
            <div className="w-8 h-8 rounded-full bg-blue-50 flex items-center justify-center mr-3">
              <span className="text-blue-600 font-medium">1</span>
            </div>
            <h2 className="text-lg font-normal text-gray-800">源Excel文件</h2>
          </div>
          
          <div 
            {...getSourceRootProps({ 
              className: 'cursor-pointer border-2 border-dashed border-gray-200 rounded-lg bg-gray-50 hover:bg-gray-100 transition-colors duration-200 py-8 px-6 flex flex-col items-center justify-center' 
            })}
          >
            <input {...getSourceInputProps()} />
            <svg xmlns="http://www.w3.org/2000/svg" className="h-10 w-10 text-gray-400 mb-3" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 13h6m-3-3v6m5 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
            {sourceFile ? (
              <div className="text-center">
                <p className="text-gray-900 font-medium truncate max-w-full">{sourceFile.name}</p>
                <p className="text-sm text-gray-500 mt-1">点击更换文件</p>
              </div>
            ) : (
              <div className="text-center">
                <p className="text-gray-600">拖放源Excel文件到这里</p>
                <p className="text-sm text-gray-500 mt-1">或点击选择文件</p>
              </div>
            )}
          </div>
          
          {sourceFile && Object.keys(sourceSheets).length > 0 && (
            <div className="space-y-4">
              <div>
                <label className="block text-sm text-gray-600 mb-2">工作表</label>
                <div className="relative">
                  <select 
                    className="w-full px-4 py-2.5 bg-white border border-gray-200 rounded-lg appearance-none focus:outline-none focus:ring-2 focus:ring-blue-100 focus:border-blue-400 transition-colors"
                    value={selectedSourceSheet}
                    onChange={handleSourceSheetChange}
                  >
                    <option value="">请选择工作表</option>
                    {Object.keys(sourceSheets).map(sheet => (
                      <option key={sheet} value={sheet}>{sheet}</option>
                    ))}
                  </select>
                  <div className="absolute inset-y-0 right-0 flex items-center px-2 pointer-events-none">
                    <svg className="w-5 h-5 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
                    </svg>
                  </div>
                </div>
              </div>
              
              {selectedSourceSheet && (
                <div>
                  <label className="block text-sm text-gray-600 mb-2">源列</label>
                  <div className="relative">
                    <select 
                      className="w-full px-4 py-2.5 bg-white border border-gray-200 rounded-lg appearance-none focus:outline-none focus:ring-2 focus:ring-blue-100 focus:border-blue-400 transition-colors"
                      value={selectedSourceColumn}
                      onChange={(e) => setSelectedSourceColumn(e.target.value)}
                    >
                      <option value="">请选择列</option>
                      {sourceSheets[selectedSourceSheet]?.map((column, index) => (
                        <option key={index} value={column}>{column}</option>
                      )) || []}
                    </select>
                    <div className="absolute inset-y-0 right-0 flex items-center px-2 pointer-events-none">
                      <svg className="w-5 h-5 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
                      </svg>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
        
        {/* 目标文件上传区域 */}
        <div className="space-y-6">
          <div className="flex items-center mb-4">
            <div className="w-8 h-8 rounded-full bg-blue-50 flex items-center justify-center mr-3">
              <span className="text-blue-600 font-medium">2</span>
            </div>
            <h2 className="text-lg font-normal text-gray-800">目标Excel文件</h2>
          </div>
          
          <div 
            {...getTargetRootProps({ 
              className: 'cursor-pointer border-2 border-dashed border-gray-200 rounded-lg bg-gray-50 hover:bg-gray-100 transition-colors duration-200 py-8 px-6 flex flex-col items-center justify-center' 
            })}
          >
            <input {...getTargetInputProps()} />
            <svg xmlns="http://www.w3.org/2000/svg" className="h-10 w-10 text-gray-400 mb-3" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 13h6m-3-3v6m5 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
            {targetFile ? (
              <div className="text-center">
                <p className="text-gray-900 font-medium truncate max-w-full">{targetFile.name}</p>
                <p className="text-sm text-gray-500 mt-1">点击更换文件</p>
              </div>
            ) : (
              <div className="text-center">
                <p className="text-gray-600">拖放目标Excel文件到这里</p>
                <p className="text-sm text-gray-500 mt-1">或点击选择文件</p>
              </div>
            )}
          </div>
          
          {targetFile && Object.keys(targetSheets).length > 0 && (
            <div className="space-y-4">
              <div>
                <label className="block text-sm text-gray-600 mb-2">工作表</label>
                <div className="relative">
                  <select 
                    className="w-full px-4 py-2.5 bg-white border border-gray-200 rounded-lg appearance-none focus:outline-none focus:ring-2 focus:ring-blue-100 focus:border-blue-400 transition-colors"
                    value={selectedTargetSheet}
                    onChange={handleTargetSheetChange}
                  >
                    <option value="">请选择工作表</option>
                    {Object.keys(targetSheets).map(sheet => (
                      <option key={sheet} value={sheet}>{sheet}</option>
                    ))}
                  </select>
                  <div className="absolute inset-y-0 right-0 flex items-center px-2 pointer-events-none">
                    <svg className="w-5 h-5 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
                    </svg>
                  </div>
                </div>
              </div>
              
              {selectedTargetSheet && (
                <div>
                  <label className="block text-sm text-gray-600 mb-2">目标列</label>
                  <div className="relative">
                    <select 
                      className="w-full px-4 py-2.5 bg-white border border-gray-200 rounded-lg appearance-none focus:outline-none focus:ring-2 focus:ring-blue-100 focus:border-blue-400 transition-colors"
                      value={selectedTargetColumn}
                      onChange={(e) => setSelectedTargetColumn(e.target.value)}
                    >
                      <option value="">请选择列</option>
                      {targetSheets[selectedTargetSheet]?.map((column, index) => (
                        <option key={index} value={column}>{column}</option>
                      )) || []}
                    </select>
                    <div className="absolute inset-y-0 right-0 flex items-center px-2 pointer-events-none">
                      <svg className="w-5 h-5 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
                      </svg>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      </div>
      
      {/* 处理按钮 */}
      <div className="flex justify-center mb-10">
        <button
          className="px-8 py-3 bg-blue-600 text-white font-normal rounded-lg hover:bg-blue-700 focus:outline-none focus:ring-4 focus:ring-blue-200 disabled:opacity-50 disabled:cursor-not-allowed transition-colors duration-200"
          onClick={processExcelFiles}
          disabled={!sourceFile || !targetFile || !selectedSourceSheet || !selectedTargetSheet || 
                   !selectedSourceColumn || !selectedTargetColumn || processing}
        >
          {processing ? (
            <span className="flex items-center">
              <svg className="animate-spin -ml-1 mr-2 h-5 w-5 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
              </svg>
              处理中...
            </span>
          ) : '处理文件'}
        </button>
      </div>
      
      {/* 消息区域 */}
      <div className="max-w-3xl mx-auto space-y-4">
        {/* 错误信息 */}
        {error && (
          <div className="bg-red-50 rounded-lg px-5 py-4 flex items-start">
            <svg className="h-5 w-5 text-red-500 mt-0.5 mr-3 flex-shrink-0" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20" fill="currentColor">
              <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd" />
            </svg>
            <p className="text-red-700">{error}</p>
          </div>
        )}
        
        {/* 调试信息 */}
        {debug && (
          <div className="bg-blue-50 rounded-lg px-5 py-4 flex items-start">
            <svg className="h-5 w-5 text-blue-500 mt-0.5 mr-3 flex-shrink-0" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20" fill="currentColor">
              <path fillRule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2h-1V9a1 1 0 00-1-1z" clipRule="evenodd" />
            </svg>
            <p className="text-blue-700">{debug}</p>
          </div>
        )}
        
        {/* 结果文件下载链接 */}
        {resultFile && (
          <div className="bg-green-50 rounded-lg px-5 py-4 flex items-start">
            <svg className="h-5 w-5 text-green-500 mt-0.5 mr-3 flex-shrink-0" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20" fill="currentColor">
              <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
            </svg>
            <div>
              <p className="text-green-700">
                文件处理成功！
                <a 
                  href={resultFile} 
                  download="processed_excel.xlsx"
                  className="font-medium underline ml-2 hover:text-green-800"
                >
                  点击下载处理后的文件
                </a>
              </p>
            </div>
          </div>
        )}
      </div>
      
      <div className="mt-16 pt-6 border-t border-gray-100 text-center text-sm text-gray-400">
        Excel列复制工具
      </div>
    </div>
  )
} 