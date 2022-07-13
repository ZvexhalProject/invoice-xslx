<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:fn="http://www.w3.org/2005/xpath-functions" xmlns:fn_ma="http://test.de"
    xmlns:xs="http://www.w3.org/2001/XMLSchema">
    <xsl:template match="/">
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3"
            xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
            xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
            xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
            xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"
            xr:uid="{00000000-0001-0000-0000-000000000000}">
            <sheetFormatPr baseColWidth="10" defaultColWidth="17.5703125" defaultRowHeight="24.95" customHeight="1"
                x14ac:dyDescent="0.4" />
            <cols>
                <col min="1" max="16384" width="17.5703125" style="1" />
            </cols>
            <sheetData>
                <xsl:for-each select="Rechnungskatalog/Rechnung">
                    <xsl:variable name="parent-pos" select="position()-1" />
                    <xsl:variable name="child-firstpos" select="(($parent-pos)*28)+14" />
                    <xsl:variable name="child-lastpos" select="(($parent-pos)*28)+count(Produkte/Produkt)+13" />
                    <row spans="1:5" ht="24.95" customHeight="1" x14ac:dyDescent="0.4">
                        <xsl:attribute name="r">
                            <xsl:value-of select="((position()-1)*28)+1" />
                        </xsl:attribute>
                        <c s="11" t="inlineStr">
                            <is>
                                <t>CMP -Projekt Nr.4: Office</t>
                            </is>
                        </c>
                        <c s="11"></c>
                        <c s="11"></c>
                        <c s="11" t="inlineStr">
                            <is>
                                <t>RECHNUNG</t>
                            </is>
                        </c>
                    </row>
                    <row spans="1:5" ht="24.95" customHeight="1" x14ac:dyDescent="0.4">
                        <c s="2" />
                        <c s="2" />
                        <c s="2" />
                        <c s="2" />
                        <c s="2" />
                    </row>
                    <row spans="1:5" ht="24.95" customHeight="1" x14ac:dyDescent="0.5">
                        <xsl:attribute name="r">
                            <xsl:value-of select="((position()-1)*28)+4" />
                        </xsl:attribute>
                        <c s="3" t="inlineStr">
                            <is>
                                <t>Rechnungsnummer</t>
                            </is>
                        </c>
                        <c s="4" />
                        <c s="5" t="inlineStr">
                            <is>
                                <t>
                                    <xsl:value-of select="Rechnungsnummer" />
                                </t>
                            </is>
                        </c>
                        <c s="4" />
                        <c s="13" t="inlineStr">
                            <is>
                                <t>
                                    <xsl:value-of select="Unternehmensname" />
                                </t>
                            </is>
                        </c>
                    </row>
                    <row spans="1:5" ht="24.95" customHeight="1" x14ac:dyDescent="0.5">
                        <c s="3" t="inlineStr">
                            <is>
                                <t>Kundennummer</t>
                            </is>
                        </c>
                        <c s="4" />
                        <c s="5" t="inlineStr">
                            <is>
                                <t>
                                    <xsl:value-of select="Kundennummer" />
                                </t>
                            </is>
                        </c>
                        <c s="4" />
                        <c s="14" t="inlineStr">
                            <is>
                                <t>
                                    <xsl:value-of select="concat(Unternehmensadresse,' ',Unternehmensadressennummer)" />
                                </t>
                            </is>
                        </c>
                    </row>
                    <row spans="1:5" ht="24.95" customHeight="1" x14ac:dyDescent="0.5">
                        <c s="3" t="inlineStr">
                            <is>
                                <t>Datum</t>
                            </is>
                        </c>
                        <c s="4" />
                        <c s="6">
                            <v>
                                <xsl:value-of select="Kaufdatum" />
                            </v>
                        </c>
                        <c s="4" />
                        <c s="14" t="inlineStr">
                            <is>
                                <t>
                                    <xsl:value-of select="Unternehmensort" />
                                </t>
                            </is>
                        </c>
                    </row>
                    <row spans="1:5" ht="24.95" customHeight="1" x14ac:dyDescent="0.5">
                        <c s="7" />
                        <c s="4" />
                        <c s="4" />
                        <c s="4" />
                        <c s="14" t="inlineStr">
                            <is>
                                <t>
                                    <xsl:value-of select="Unternehmensland" />
                                </t>
                            </is>
                        </c>
                    </row>
                    <row spans="1:5" ht="24.95" customHeight="1" x14ac:dyDescent="0.4">
                        <c s="3" t="inlineStr">
                            <is>
                                <t>
                                    <xsl:value-of select="concat(Kundenvorname,' ',Kundennachname)" />
                                </t>
                            </is>
                        </c>
                    </row>
                    <row spans="1:5" ht="24.95" customHeight="1" x14ac:dyDescent="0.4">
                        <c s="3" t="inlineStr">
                            <is>
                                <t>
                                    <xsl:value-of select="concat(Kundenstrasse,' ',Kundenstrassennummer)" />
                                </t>
                            </is>
                        </c>
                    </row>
                    <row spans="1:5" ht="24.95" customHeight="1" x14ac:dyDescent="0.4">
                        <c s="3" t="inlineStr">
                            <is>
                                <t>
                                    <xsl:value-of select="Kundenstrassenort" />
                                </t>
                            </is>
                        </c>
                    </row>
                    <row spans="1:5" ht="24.95" customHeight="1" x14ac:dyDescent="0.4">
                        <c s="3" t="inlineStr">
                            <is>
                                <t>
                                    <xsl:value-of select="Kundenland"/>
                                </t>
                            </is>
                        </c>
                    </row>
                    <row spans="1:5" ht="24.95" customHeight="1" x14ac:dyDescent="0.4">
                        <xsl:attribute name="r">
                            <xsl:value-of select="((position()-1)*28)+13" />
                        </xsl:attribute>
                        <c s="15" t="inlineStr">
                            <is>
                                <t>MENGE</t>
                            </is>
                        </c>
                        <c s="16" t="inlineStr">
                            <is>
                                <t>ARTIKELNUMMER</t>
                            </is>
                        </c>
                        <c s="16" t="inlineStr">
                            <is>
                                <t>ARTIKELNAME</t>
                            </is>
                        </c>
                        <c s="15" t="inlineStr">
                            <is>
                                <t>EINZELPREIS</t>
                            </is>
                        </c>
                        <c s="18" t="inlineStr">
                            <is>
                                <t>BETRAG</t>
                            </is>
                        </c>
                    </row>
                    <xsl:for-each select="Produkte/Produkt">
                        <row spans="1:5" ht="24.95" customHeight="1" x14ac:dyDescent="0.4">
                            <xsl:attribute name="r">
                                <xsl:value-of select="(($parent-pos)*28)+position()+13" />
                            </xsl:attribute>
                            <c s="8">
                                <v>
                                    <xsl:value-of select="Menge" />
                                </v>
                            </c>
                            <c s="8" t="inlineStr">
                                <is>
                                    <t>
                                        <xsl:value-of select="Artikelnummer" />
                                    </t>
                                </is>
                            </c>
                            <c s="8" t="inlineStr">
                                <is>
                                    <t>
                                        <xsl:value-of select="Artikelname" />
                                    </t>
                                </is>
                            </c>
                            <c s="12">
                               <v>
                                        <xsl:value-of select="Preis" />
                               </v>
                            </c>
                            <c s="17">
                                <f>
                                    <xsl:value-of
                                        select="concat('A',((($parent-pos)*28)+13+position()),'*D',((($parent-pos)*28)+13+position()))" />
                                </f>
                                <!-- <xsl:value-of
                                        select="concat('A',((count(../preceding-sibling::Rechnung)+1)*28)+14+position(),'*D',((count(../preceding-sibling::Rechnung)+1)*28)+14+position())" />
                                </f> -->
                            </c>
                        </row>
                    </xsl:for-each>
                    <row spans="1:5" ht="24.95" customHeight="1" thickBot="1" x14ac:dyDescent="0.45">
                        <xsl:attribute name="r">
                            <xsl:value-of select="((position()-1)*28)+26" />
                        </xsl:attribute>
                    </row>
                    <row spans="1:5" ht="24.95" customHeight="1" thickBot="1" x14ac:dyDescent="0.45">
                        <c s="9" t="s"></c>
                        <c s="9" t="s"></c>
                        <c s="9" t="s"></c>
                        <c s="9" t="inlineStr">
                            <is>
                                <t>Summe</t>
                            </is>
                        </c>
                        <c s="19">
                            <!-- <xsl:value-of select="$child-firstpos" />
                         <xsl:value-of select="$child-lastpos" /> -->
                            <f>
                                <xsl:value-of select="concat('SUM(E',$child-firstpos,':E',$child-lastpos,')')" />
                            </f>
                        </c>
                    </row>
                </xsl:for-each>

            </sheetData>
            <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />
            <pageSetup paperSize="9" orientation="portrait" horizontalDpi="4294967293" verticalDpi="0" r:id="rId1" />
        </worksheet>
    </xsl:template>

</xsl:stylesheet>