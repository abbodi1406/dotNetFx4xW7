<xsl:stylesheet version="1.0"
            xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
            xmlns:msxsl="urn:schemas-microsoft-com:xslt"
            exclude-result-prefixes="msxsl"
            xmlns:wix="http://schemas.microsoft.com/wix/2006/wi"
            xmlns:my="my:my">
    <xsl:output method="xml" indent="yes" />
    <xsl:strip-space elements="*"/>
    <xsl:template match="@*|node()">
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
        </xsl:copy>
    </xsl:template>
    <xsl:template match="@EmbedCab[parent::wix:Media[parent::wix:Product[parent::wix:Wix]]]">
        <xsl:attribute name="EmbedCab">
            <xsl:value-of select="'no'"/>
        </xsl:attribute>
    </xsl:template>
</xsl:stylesheet>