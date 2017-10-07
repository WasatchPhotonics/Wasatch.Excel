Dim spectrometer As WasatchNET.spectrometer
Dim pixels As Integer

Public Sub buttonInitialize_Click()
    ' VBA/VB6 can't use .NET static methods, so access through wrapper
    Dim wrapper As New WasatchNET.DriverVBAWrapper
    Dim driver As WasatchNET.driver
    Set driver = wrapper.instance
    
    Dim numberOfSpectrometers As Integer
    numberOfSpectrometers = driver.openAllSpectrometers()
    If (numberOfSpectrometers <= 0) Then
        MsgBox "No spectrometers found"
        Return
    End If
    
    Set spectrometer = driver.getSpectrometer(0)
    pixels = spectrometer.pixels
    
    Dim modelConfig As WasatchNET.modelConfig
    Set modelConfig = spectrometer.modelConfig
        
    Range("model") = spectrometer.model
    Range("serialNumber") = spectrometer.serialNumber
    Range("pixels") = pixels
    
    Dim wavelengths() As Double
    Dim wavenumbers() As Double
    Dim excitationNM As Integer
    wavelengths = spectrometer.wavelengths
    wavenumbers = spectrometer.wavenumbers
    excitationNM = modelConfig.excitationNM
    
    Dim wavecalCoeffs() As Single
    wavecalCoeffs = modelConfig.wavecalCoeffs
    Range("coeff0") = wavecalCoeffs(0)
    Range("coeff1") = wavecalCoeffs(1)
    Range("coeff2") = wavecalCoeffs(2)
    Range("coeff3") = wavecalCoeffs(3)
    
    rowStart = Range("pixelHeader").Row + 1
    colPixel = Range("pixelHeader").Column
    colWavelength = Range("wavelengthHeader").Column
    colWavenumber = Range("wavenumberHeader").Column
    For i = 0 To pixels - 1
        ActiveSheet.Cells(rowStart + i, colPixel).Value = i
        ActiveSheet.Cells(rowStart + i, colWavelength).Value = wavelengths(i)
        If excitationNM > 0 Then
            ActiveSheet.Cells(rowStart + i, colWavenumber).Value = wavenumbers(i)
        End If
    Next i
End Sub

Public Sub buttonAcquire_Click()
    If spectrometer Is Nothing Then
        MsgBox "Please initialize spectrometer first"
        Return
    End If
    
    spectrometer.integrationTimeMS = Range("integTimeMS")
    spectrometer.scanAveraging = Range("scansToAverage")
    spectrometer.boxcarHalfWidth = Range("boxcarHalfWidth")
    
    Dim spectrum() As Double
    spectrum = spectrometer.getSpectrum()
    colIntensity = Range("intensityHeader").Column
    rowStart = Range("intensityHeader").Row + 1
    For i = 0 To pixels - 1
        ActiveSheet.Cells(rowStart + i, colIntensity).Value = spectrum(i)
    Next i
End Sub
