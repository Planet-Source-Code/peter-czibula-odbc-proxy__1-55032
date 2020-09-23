
The Client uses a VB class module (rstWinsock) which simulates a ADODB recordset. The following has been implemented:

- PROPERTY rstWinsock.RemoteHostIP (the remote host IP address)
- PROPERTY rstWinsock.RemoteHostPort (the remote host port)
- PROPERTY rstWinsock.TimeOutSecs (the Timeout in seconds that the class waits for a server response)
- METHOD rstWinsock.ConnectSocket (establish the connection)
- PROPERTY rstWinsock.State (Connected, Not connected, Error, Timeout)
  if the state = Connected, everything processed ok (check this first)
  if the state = Error, check the ErrCode and ErrDescr for the error returned
- METHOD rstWinsock.ConnectRemoteDSN (initialize connection to remote DSN)
  check state = Connected to see if everything is OK
- METHOD rstWinsock.ExecuteSQL (launch a SQL statement on the remote DSN)
  if the SQL is a SELECT statement, the class fetches the fields and the first record
  check state = Connected to see if everything is OK
  if state = connected, 
    - check EOF to see if records are available
    - access all the fields returned with the FIELDS collection
- METHOD rstWinsock.MoveNext (launch a move-to-next record on the remote DSN)
  check state = Connected to see if everything is OK
  if state = connected, check EOF to see if records are available  

N.B. the class does not use any reference to the winsock component (it uses the excellent custom class available on www.vbip.com, written by Oleg Gdalevich, which simulates the Winsock control)

N.B. I only implemented the following field data types from ADO (N5=adSmallInt, N10=adInteger, N30=adCurrency, C=adVarChar, D=adDate). This can be changed in the "Sub GetFields" and the "Sub MoveNext".

N.B. The client is based on Server V2, but works also on the previous version (except that the date format returned from the Server is different "01/01/2000" vs "2000/01/01")

N.B. The rstWInsock class can even be compiled into a DLL and thus could be used on a IIS/ASP server, but I do not need it for the moment.

