Attribute VB_Name = "mod_variables"
Option Explicit

'ByVal szDomainName As String, _
                                       'ByVal szOrganizationName As String, _
                                       'ByVal szAdministrativeGroupName As String, _
                                       'ByVal szUserName As String, _
                                       'ByVal szUserPwd As String, _
                                       'ByRef szServerList As String, _
                                       'ByVal szDirectoryServer
                                       
                                       
Public bHaltProcess As Boolean
                                       
Public Type Errors_Struct
    Number As String
    Desc As String
    source As String
    Meaning As String
    DateTime As String
End Type

Public Type ADUser_Struct
    MainIndex As Integer
    CN As String
    SamAccountName As String
    Domain As String
    AccountDisabled As String
    AccountExpirationDate As String
    ADsPath As String
    BadLoginAddress As String
    BadLoginCount As String
    Class As String
    Department As String
    Description As String
    Division As String
    EmailAddress As String
    EmployeeID As String
    FaxNumber As String
    FirstName As String
    FullName As String
    GraceLoginsAllowed As String
    GraceLoginsRemaining As String
    Guid As String
    HomeDirectory As String
    HomePage As String
    IsAccountLocked As String
    Languages As String
    LastFailedLogin As String
    LastLogin As String
    LastLogoff As String
    LastName As String
    LoginHours As String
    LoginScript As String
    LoginWorkstations As String
    Manager As String
    MaxLogins As String
    MaxStorage As String
    name As String
    NamePrefix As String
    NameSuffix As String
    OfficeLocations As String
    OtherName As String
    Parent As String
    PasswordExpirationDate As String
    PasswordLastChanged As String
    PasswordMinimumLength As String
    PasswordRequired As String
    Picture As String
    PostalAddresses As String
    PostalCodes As String
    Profile As String
    RequireUniquePassword As String
    Schema As String
    SeeAlso As String
    TelephoneHome As String
    TelephoneMobile As String
    TelephoneNumber As String
    TelephonePager As String
    Title As String
End Type

Public Type Domains_Struct
    name As String
    ADUsers() As ADUser_Struct
End Type
Public Type MailBoxes_Struct
    name As String
    FullName As String
    Alias As String
    FileName As String
    ForwardingStyle As String
    ForwardTo As String
    HideFromAddressBook As String
    IncomingLimit As String
    OutgoingLimit As String
    ProxyAddresses As String
    RestrictedAddresses As String
    RestrictedAddressList As String
    SMTPEmail As String
    TargetAddress As String
    Size As Double
    DisplaySize As String
    MessageCount As String
    DateCreated As String
    DateLastModified As String
    ADInfo As ADUser_Struct
    Flagged_2B_Processed As Boolean
End Type
Public Type MailboxStoreDB_Struct
    name As String
    DaysBeforeDeletedMailboxCleanup As String
    DaysBeforeGarbageCollection As String
    DBPath As String
    Enabled As String
    GarbageCollectOnlyAfterBackup As String
    OfflineAddressList As String
    OverQuotaLimit As String
    HardLimit As String
    PublicStoreDB As String
    SLVPath As String
    Status As String
    StoreQuota As String
    MBX() As MailBoxes_Struct
    MailBoxCount As Integer
End Type
Public Type StorageGroups_Struct
    name As String
    LogFilePath As String
    SystemFilePath As String
    ZeroDatabase As String
    CircularLogging As String
    FieldCount As String
    MBSDB() As MailboxStoreDB_Struct
    MailBoxStoreDBCount As Integer
End Type
Public Type Server_Struct
    name As String
    DaysBeforeLogFileRemoval As String
    DirectoryServer As String
    ExchangeVersion As Variant
    MessageTrackingEnabled As String
    SubjectLoggingEnabled As String
    ServerType As String
    SG() As StorageGroups_Struct
    StorageGroupCount As Integer
    Errors() As Errors_Struct
End Type
Public Type Exchange_Struct
    Svrs() As Server_Struct
    ServerCount As Integer
End Type


Public Type Drive_Info_Struct
    ShareName As String
    VolumeName As String
    RootFolder As String
    Path As String
    IsReady As String
    FileSystem As String
    FreeSpace As String
    TotalSpace As String
    Letter As String
    Type As String
    Model As String
    Serial As String
End Type

Public Type MailBox_Struct
    MailBoxPath As String
    MailBoxFileName As String
    Organization As String
    Group As String
    CN As String
End Type

Public Type INI_Struct
    RootPath As String
    UseRootPath As Boolean
    INIPath As String
    INIFileName As String
    EmailServer As String
    ExMergePath As String
    PSTPath As String
    LogPath As String
    Loglevel As Integer
    ExportUserData As Boolean
    ExportFolderRules As Boolean
    ExportDumpster As Boolean
    ExportFolderData As Boolean
End Type

Public Type Accounts_2B_Processed_Struct
    index As Integer
End Type

Public Type Current_Struct
    Server As String
    StorageGroup As String
    MailBoxStoreDBs As String
    MailBox As String
    Domain As String
    Organisation As String
    Container As String
    Object As String
End Type

Public Type DateRange_Struct
    Requestor As String
    InputDate As String
    ReturnDate As String
End Type

Public Type SysInfo_Struct
    ComputerName As String
    UserName As String
    Version As String
    AD_Exists As Boolean
    Exchange_Exists As Boolean
End Type

Public Type Main_Struct
    Domains() As Domains_Struct
    Drives() As Drive_Info_Struct
    MailBoxFile As MailBox_Struct
    INIFile As INI_Struct
    Accounts() As Accounts_2B_Processed_Struct
    Exch As Exchange_Struct
    Current As Current_Struct
    DateRange As DateRange_Struct
    Sys As SysInfo_Struct
End Type


Public Main As Main_Struct
Public Servers() As Server_Struct
Public CurrentMailBoxes() As MailBoxes_Struct
