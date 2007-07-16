from System.Net import WebRequest
from System.IO import StreamReader
from System.Text import Encoding
from System import DateTime

URI = 'http://www.oanda.com/convert/fxhistory'
PARAMETERS = "date1=%s&date=%s&lang=en&result=1&format=CSV&date_fmt=normal&exch=%s&expr=%s" 

def GetRealTimeData(date1, date2, exch1, exch2):
    parameters = PARAMETERS % (date1, date2, exch1, exch2)
    fullURI = URI + "/" + parameters
    print 'Starting Reguest for:', fullURI
    start = DateTime.Now
    result = FetchURIWithParameters(URI, parameters)
    print 'Finished. Took %s seconds' % ((DateTime.Now - start).TotalMilliseconds / 1000.0)
    
    # scrape the results from the returned web page
    indexOfOpenPRE = result.find('<PRE>')
    indexOfClosePRE = result.find('</PRE>')
    return result[indexOfOpenPRE+5:indexOfClosePRE], fullURI


def FetchURIWithParameters(uri, parameters=None):
    request = WebRequest.Create(uri)
    if parameters is not None:
        request.ContentType = "application/x-www-form-urlencoded"
        request.Method = "POST"
        bytes = Encoding.ASCII.GetBytes(parameters)
        request.ContentLength = bytes.Length
        reqStream = request.GetRequestStream()
        reqStream.Write(bytes, 0, bytes.Length)
        reqStream.Close()

    response = request.GetResponse()
    result = StreamReader(response.GetResponseStream()).ReadToEnd()
    return result

