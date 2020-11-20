package microsoft.exchange.webservices.data.search;

public enum MailboxSearchLocation {
  /// <summary>
  /// Primary only (Exchange 2013 or later).
  /// </summary>
  PrimaryOnly,

  /// <summary>
  /// Archive only (Exchange 2013 or later).
  /// </summary>
  ArchiveOnly,

  /// <summary>
  /// Both Primary and Archive (Exchange 2013 or later).
  /// </summary>
  All,
}
