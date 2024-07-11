package com.github.pjfanning.xlsx;

public enum CommentsImplementationType {
  /**
   * The default <code>CommentsTable</code> in POI
   */
  POI_DEFAULT,
  /**
   * The temp file backed <code>CommentsTable</code> in <code>poi-shared-strings</code>.
   * Saves on memory but still has good performance, especially if full-format text is set to false.
   * <p>
   *     poi-shared-strings is an optional dependency since v5.0.0 - you will need to include it in your project if you want to use this implementation
   * </p>
   * @see StreamingReader.Builder#setFullFormatRichText(boolean)
   * @see StreamingReader.Builder#setEncryptCommentsTempFile(boolean)
   */
  TEMP_FILE_BACKED,
  /**
   * The concurrent map backed <code>CommentsTable</code> in <code>poi-shared-strings</code>.
   * More performant if full-format text is set to false.
   * <p>
   *     poi-shared-strings is an optional dependency since v5.0.0 - you will need to include it in your project if you want to use this implementation
   * </p>
   * @see StreamingReader.Builder#setFullFormatRichText(boolean)
   */
  CUSTOM_MAP_BACKED
}
