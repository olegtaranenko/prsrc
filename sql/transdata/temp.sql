SELECT sGuideNomenk.*, sGuideFormuls.Formula, sGuideSource.sourceName, sGuideFormuls_1.Formula AS formulaW FROM sGuideFormuls AS sGuideFormuls_1 INNER JOIN ((sGuideNomenk INNER JOIN sGuideSource ON sGuideNomenk.sourId = sGuideSource.sourceId) INNER JOIN sGuideFormuls ON sGuideNomenk.formulaNom = sGuideFormuls.nomer) ON sGuideFormuls_1.nomer = sGuideNomenk.formulaNomW WHERE (((sGuideNomenk.klassId)=69)) ORDER BY sGuideNomenk.nomNom ;