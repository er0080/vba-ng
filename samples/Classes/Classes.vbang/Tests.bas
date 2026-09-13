Attribute VB_Name = "Tests"
Option Explicit

' Test procedures for the class modules: lifetime, properties, default members,
' For Each, Implements, events, late binding, and the predeclared instance.

'@Test
Public Sub NewLedger_StartsEmpty()
    Dim book As Ledger
    Set book = New Ledger
    Assert.AreEqual 0, book.Count
    Assert.AreEqual "ledger", book.Name
End Sub

'@Test
Public Sub Post_AddsAnEntryTheDefaultMemberReads()
    Dim book As New Ledger
    book.Post 12.5
    book.Post 20
    Assert.AreEqual 2, book.Count
    Assert.AreEqual 12.5, book(1)
    Assert.AreEqual 20, book.Entry(2)
End Sub

'@Test
Public Sub ForEach_WalksTheEntries()
    Dim book As New Ledger
    Dim entry As Variant
    Dim total As Currency
    book.Post 1
    book.Post 2
    book.Post 3
    For Each entry In book
        total = total + entry
    Next
    Assert.AreEqual 6, total
End Sub

'@Test
Public Sub Implements_ReachesTheImplementation()
    Dim book As Ledger
    Dim account As IAccount
    Set book = OpenLedger("main")
    book.Post 40
    book.Post 2
    Set account = book
    Assert.AreEqual "ledger main", account.Label
    Assert.AreEqual 42, account.Balance
    Assert.AreEqual 42, TotalOf(book)
    Assert.IsTrue TypeOf account Is Ledger
    Assert.IsTrue account Is book
End Sub

'@Test
Public Sub Events_RunTheHandlerAndCanCancel()
    Dim book As New Ledger
    Dim watcher As New Auditor
    Set watcher.Target = book
    watcher.Limit = 100
    book.Post 30
    book.Post 250
    Assert.AreEqual 2, watcher.Seen
    Assert.AreEqual 1, book.Count
    Assert.AreEqual 30, book.Entry(1)
    Set watcher.Target = Nothing
    book.Post 5
    Assert.AreEqual 2, watcher.Seen
End Sub

'@Test
Public Sub Terminate_RunsWhenTheLastReferenceGoes()
    Dim before As Long
    before = Live
    Dim held As Tracker
    Set held = New Tracker
    Assert.AreEqual before + 1, Live
    Dim second As Tracker
    Set second = held
    Set held = Nothing
    Assert.AreEqual before + 1, Live
    Set second = Nothing
    Assert.AreEqual before, Live
End Sub

'@Test
Public Sub Terminate_RunsWhenALocalGoesOutOfScope()
    Dim before As Long
    before = Live
    MakeAndDrop
    Assert.AreEqual before, Live
End Sub

'@Test
Public Sub PredeclaredInstance_IsReachedByTheClassName()
    Settings.Unit = "EUR"
    Assert.AreEqual "EUR", Settings.Unit
    Dim other As Settings
    Set other = New Settings
    other.Unit = "USD"
    Assert.AreEqual "EUR", Settings.Unit
    Assert.IsFalse other Is Settings
End Sub

'@Test
Public Sub LateBinding_ReachesAClassByName()
    Dim book As Object
    Set book = New Ledger
    book.Post 7
    Assert.AreEqual 1, book.Count
    Assert.AreEqual 7, CallByName(book, "Entry", VbGet, 1)
    Assert.AreEqual "Ledger", DescribeAny(book)
End Sub

'@Test
Public Sub Collection_HoldsAndReleasesObjects()
    Dim before As Long
    before = Live
    Dim held As Collection
    Set held = New Collection
    held.Add New Tracker
    held.Add New Tracker
    Assert.AreEqual before + 2, Live
    held.Remove 1
    Assert.AreEqual before + 1, Live
    Set held = Nothing
    Assert.AreEqual before, Live
End Sub

Private Sub MakeAndDrop()
    Dim scratch As Tracker
    Set scratch = New Tracker
End Sub
