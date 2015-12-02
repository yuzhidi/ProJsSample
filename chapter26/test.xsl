<?xml version='1.0'?>
<xsl:stylesheet version="1.0" 
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
   <xsl:template match="/">
      <html>
      <head>
      <title>XSL Test</title>
      </head>
      <body>
         <xsl:for-each select="example/demo">
          <h1><xsl:value-of select="."/></h1>
         </xsl:for-each>
      </body>
      </html>
   </xsl:template>
</xsl:stylesheet>
