<?xml version="1.0" encoding="utf-8"?>

<!-- Deployment Wizard Build : 13.0.0-B20171108 -->
<jnlp codebase="https://hod.serpro.gov.br/hod/" href="">
<!-- <jnlp codebase="https://hod.serpro.gov.br/hod/" href="https://hod.serpro.gov.br/a83016cv/hodcivws/hodcivws.jnlp"> -->
  <information>
    <title>Painel de controle</title>
    <vendor>IBM Corporation</vendor>
    <description>Host On-Demand</description>
    <icon href="images/hodSplash.png" kind="splash"/>
    <icon href="images/hodIcon.png" kind="shortcut"/>
    <icon href="images/hodIcon.png" kind="default"/>
    <offline-allowed/>
    <shortcut online="false">
    </shortcut>
  </information>
  <security>
    <all-permissions/>
  </security>
  <resources>
    <j2se version="1.3+"/>
    <jar href="WSCachedSupporter2.jar" download="eager" main="true"/>
    <jar href="CachedAppletInstaller2.jar" download="eager"/>
    <property name="jnlp.hod.TrustedJNLP" value="true"/>
    <property name="jnlp.hod.WSFrameTitle" value="Painel$SPACE$de$SPACE$controle"/>
    <property name="jnlp.hod.DocumentBase" value="https://hod.serpro.gov.br/a83016cv/hodcivws/hodcivws.jnlp"/>
    <property name="jnlp.hod.PreloadComponentList" value="HABASE;HODBASE;HODIMG;HACP;HAFNTIB;HAFNTAP;HA5250X;HA3270;HODCUT;HAMACUI;HODCFG;HAVT;HA5250P;HASQL;HODTOIA;HAPD3270;HA5250E;HAKEYMP;HA3270X;HODPOPPAD;HACOLOR;HAKEYPD;HA3270P;HASSL;HODHLL;HA5250;HODDBAS;HASSH;HASSLITE;HODSQL;HODMAC;HODTLBR;HAFTP;HODZP;HACICS;HAHOSTG;HAPRINT;HASLP;HACLTAU;HODAPPL;HAMACRT;HODSSL;HAXFER"/>
    <property name="jnlp.hod.DebugComponents" value="false"/>
    <property name="jnlp.hod.DebugCachedClient" value="false"/>
    <property name="jnlp.hod.UpgradePromptResponse" value="Prompt"/>
    <property name="jnlp.hod.UpgradePercent" value="100"/>
    <property name="jnlp.hod.InstallerFrameWidth" value="550"/>
    <property name="jnlp.hod.InstallerFrameHeight" value="300"/>
    <property name="jnlp.hod.ParameterFile" value="HODData/hodcivws/params.txt"/>
    <property name="jnlp.hod.UserDefinedParameterFile" value="HODData/hodcivws/udparams.txt"/>
    <property name="jnlp.hod.CachedClientSupportedApplet" value="com.ibm.eNetwork.HOD.HostOnDemand"/>
    <property name="jnlp.hod.AdditionalArchives" value="Print_Ass"/>
    <property name="jnlp.hod.CachedClient" value="true"/>

    <!--Para luname via post -->
    

    <!-- Customizacoes para HOD Serpro -->
    
    <property name="jnlp.hod.EnableHTMLOverrides" value="true"/>
    <property name="jnlp.hod.TargetedSessionList" value="Terminal 3270"/>
    <property name="jnlp.hod.LUName" value="Terminal 3270=AWVAD4VS"/> 

  </resources>
  <application-desc main-class="com.ibm.eNetwork.HOD.cached.wssupport.WSCachedSupporter"/>
</jnlp>
