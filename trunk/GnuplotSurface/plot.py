from System.Diagnostics import Process


def GenerateGnuplotData(sheet):
    out = []
    for y in range(1, sheet.MaxRow + 1):
        row = []
        for x in range(1, sheet.MaxCol + 1):
            cell = sheet[x, y]
            value = cell.Value
            u = (6.0 / 50) * x - 3
            v = (6.0 / 50) * y - 3
            row.append('%s %s %s' % (u, v, value))
        out.append(row)
    return '\n\n    '.join('\n    '.join(row) for row in out)


def CreateScript(template, scriptPath, data, imagePath):
    h = open(scriptPath, 'w')
    h.write(template % (imagePath.replace('\\', '/'), data))
    h.close()


def LaunchGnuplot(gnuplotPath, scriptPath, fontPath):
    proc = Process()
    proc.StartInfo.FileName = gnuplotPath
    proc.StartInfo.Arguments = '"%s"' % scriptPath
    proc.StartInfo.EnvironmentVariables['GDFONTPATH'] = fontPath
    proc.StartInfo.EnvironmentVariables['GNUPLOT_FONTPATH'] = fontPath
    proc.StartInfo.UseShellExecute = False
    proc.Start()
    proc.WaitForExit()


