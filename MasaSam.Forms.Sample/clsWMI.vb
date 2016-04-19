'using System;
Imports System.Management
Imports System.IO
'Imports Scripting
Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices


Public Class clsWMI
    Private objOS As ManagementObjectSearcher
    Private objCS As ManagementObjectSearcher
    Private objMgmt As ManagementObject
    Private m_strComputerName As String
    Private m_strManufacturer As String
    Private m_StrModel As String
    Private m_strOSName As String
    Private m_strOSVersion As String
    Private m_strSystemType As String
    Private m_strTPM As String
    Private m_strWindowsDir As String


    Public Sub New()

        objOS = New ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem")
        objCS = New ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem")
        For Each objMgmt In objOS.Get


            m_strOSName = splitTextTo("|", objMgmt("name").ToString())
            m_strOSVersion = objMgmt("version").ToString()
            m_strComputerName = objMgmt("csname").ToString()
            m_strWindowsDir = objMgmt("windowsdirectory").ToString()
        Next

        For Each objMgmt In objCS.Get
            m_strManufacturer = objMgmt("manufacturer").ToString()
            m_StrModel = objMgmt("model").ToString()
            m_strSystemType = objMgmt("systemtype").ToString
            m_strTPM = ConvertSize(objMgmt("totalphysicalmemory").ToString())
        Next
    End Sub

    '' Get Just First String
    Public Function splitTextTo(Symbol As String, Text As String)
        Dim result As String() = Text.Split(New String() {Symbol}, StringSplitOptions.None)
        For Each s As String In result
            Return s.Trim
        Next
    End Function

    '' Convert Size
    Public Function ConvertSize(Size)
        Size = CSng(Replace(Size, ",", ""))

        If Not VarType(Size) = vbSingle Then
            ConvertSize = "SIZE INPUT ERROR"
            Exit Function
        End If

        Dim suffix As String = " Bytes"
        If Size >= 1024 Then suffix = " KB"
        If Size >= 1048576 Then suffix = " MB"
        If Size >= 1073741824 Then suffix = " GB"
        If Size >= 1099511627776 Then suffix = " TB"

        Select Case suffix
            Case " KB"
                Size = Math.Round(Size / 1024, 1)
            Case " MB"
                Size = Math.Round(Size / 1048576, 1)
            Case " GB"
                Size = Math.Round(Size / 1073741824, 1)
            Case " TB"
                Size = Math.Round(Size / 1099511627776, 1)

        End Select

        ConvertSize = Size & suffix
    End Function

    Public ReadOnly Property ComputerName()
        Get
            ComputerName = m_strComputerName
        End Get

    End Property
    Public ReadOnly Property Manufacturer()
        Get
            Manufacturer = m_strManufacturer
        End Get

    End Property
    Public ReadOnly Property Model()
        Get
            Model = m_StrModel
        End Get

    End Property
    Public ReadOnly Property OsName()
        Get
            OsName = m_strOSName
        End Get

    End Property

    Public ReadOnly Property OSVersion()
        Get
            OSVersion = m_strOSVersion
        End Get

    End Property
    Public ReadOnly Property SystemType()
        Get
            SystemType = m_strSystemType
        End Get

    End Property
    Public ReadOnly Property TotalPhysicalMemory()
        Get
            TotalPhysicalMemory = m_strTPM
        End Get

    End Property

    Public ReadOnly Property WindowsDirectory()
        Get
            WindowsDirectory = m_strWindowsDir
        End Get

    End Property

End Class
