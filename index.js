const parse5 = require('parse5');

const canvas = "<div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.0\" data-sp-controldata=\"&#123;&quot;controlType&quot;&#58;3,&quot;displayMode&quot;&#58;2,&quot;id&quot;&#58;&quot;cec5a8b9-1d58-41a3-980c-d96793a6e5ce&quot;,&quot;position&quot;&#58;&#123;&quot;zoneIndex&quot;&#58;1,&quot;sectionIndex&quot;&#58;1,&quot;controlIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;8&#125;,&quot;webPartId&quot;&#58;&quot;daf0b71c-6de8-4ef7-b511-faae7c388708&quot;&#125;\"><div data-sp-webpart=\"\" data-sp-webpartdataversion=\"2.1\" data-sp-webpartdata=\"&#123;&quot;id&quot;&#58;&quot;daf0b71c-6de8-4ef7-b511-faae7c388708&quot;,&quot;instanceId&quot;&#58;&quot;cec5a8b9-1d58-41a3-980c-d96793a6e5ce&quot;,&quot;title&quot;&#58;&quot;Highlighted content&quot;,&quot;description&quot;&#58;&quot;Dynamically display content based on location, type, and filtering.&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&quot;title&quot;&#58;&quot;Most recent documents&quot;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&quot;baseUrl&quot;&#58;&quot;https&#58;//m365x324230.sharepoint.com/sites/mtest_001&quot;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;2.1&quot;,&quot;properties&quot;&#58;&#123;&quot;displayMaps&quot;&#58;&#123;&quot;1&quot;&#58;&#123;&quot;headingText&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;SiteTitle&quot;]&#125;,&quot;headingUrl&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;SitePath&quot;]&#125;,&quot;title&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;UserName&quot;,&quot;Title&quot;]&#125;,&quot;personImageUrl&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;ProfileImageSrc&quot;]&#125;,&quot;name&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;Name&quot;]&#125;,&quot;initials&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;Initials&quot;]&#125;,&quot;itemUrl&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;WebPath&quot;]&#125;,&quot;activity&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;ModifiedDate&quot;]&#125;,&quot;previewUrl&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;PreviewUrl&quot;,&quot;PictureThumbnailURL&quot;]&#125;,&quot;iconUrl&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;IconUrl&quot;]&#125;,&quot;accentColor&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;AccentColor&quot;]&#125;,&quot;cardType&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;CardType&quot;]&#125;,&quot;tipActionLabel&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;TipActionLabel&quot;]&#125;,&quot;tipActionButtonIcon&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;TipActionButtonIcon&quot;]&#125;&#125;,&quot;2&quot;&#58;&#123;&quot;column1&quot;&#58;&#123;&quot;heading&quot;&#58;&quot;&quot;,&quot;sources&quot;&#58;[&quot;FileExtension&quot;],&quot;width&quot;&#58;34&#125;,&quot;column2&quot;&#58;&#123;&quot;heading&quot;&#58;&quot;Title&quot;,&quot;sources&quot;&#58;[&quot;Title&quot;],&quot;linkUrls&quot;&#58;[&quot;WebPath&quot;],&quot;width&quot;&#58;250&#125;,&quot;column3&quot;&#58;&#123;&quot;heading&quot;&#58;&quot;Modified&quot;,&quot;sources&quot;&#58;[&quot;ModifiedDate&quot;],&quot;width&quot;&#58;100&#125;,&quot;column4&quot;&#58;&#123;&quot;heading&quot;&#58;&quot;Modified By&quot;,&quot;sources&quot;&#58;[&quot;Name&quot;],&quot;width&quot;&#58;150&#125;&#125;,&quot;3&quot;&#58;&#123;&quot;id&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;UniqueID&quot;]&#125;,&quot;edit&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;edit&quot;]&#125;,&quot;DefaultEncodingURL&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;DefaultEncodingURL&quot;]&#125;,&quot;FileExtension&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;FileExtension&quot;]&#125;,&quot;FileType&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;FileType&quot;]&#125;,&quot;Path&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;Path&quot;]&#125;,&quot;PictureThumbnailURL&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;PictureThumbnailURL&quot;]&#125;,&quot;SiteID&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;SiteID&quot;]&#125;,&quot;SiteTitle&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;SiteTitle&quot;]&#125;,&quot;Title&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;Title&quot;]&#125;,&quot;UniqueID&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;UniqueID&quot;]&#125;,&quot;WebId&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;WebId&quot;]&#125;,&quot;WebPath&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;WebPath&quot;]&#125;&#125;,&quot;4&quot;&#58;&#123;&quot;headingText&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;SiteTitle&quot;]&#125;,&quot;headingUrl&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;SitePath&quot;]&#125;,&quot;title&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;UserName&quot;,&quot;Title&quot;]&#125;,&quot;personImageUrl&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;ProfileImageSrc&quot;]&#125;,&quot;name&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;Name&quot;]&#125;,&quot;initials&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;Initials&quot;]&#125;,&quot;itemUrl&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;WebPath&quot;]&#125;,&quot;activity&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;ModifiedDate&quot;]&#125;,&quot;previewUrl&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;PreviewUrl&quot;,&quot;PictureThumbnailURL&quot;]&#125;,&quot;iconUrl&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;IconUrl&quot;]&#125;,&quot;accentColor&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;AccentColor&quot;]&#125;,&quot;cardType&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;CardType&quot;]&#125;,&quot;tipActionLabel&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;TipActionLabel&quot;]&#125;,&quot;tipActionButtonIcon&quot;&#58;&#123;&quot;sources&quot;&#58;[&quot;TipActionButtonIcon&quot;]&#125;&#125;&#125;,&quot;query&quot;&#58;&#123;&quot;contentLocation&quot;&#58;1,&quot;contentTypes&quot;&#58;[1],&quot;sortType&quot;&#58;1,&quot;filters&quot;&#58;[&#123;&quot;filterType&quot;&#58;1,&quot;value&quot;&#58;&quot;&quot;&#125;],&quot;documentTypes&quot;&#58;[99],&quot;advancedQueryText&quot;&#58;&quot;&quot;&#125;,&quot;templateId&quot;&#58;1,&quot;maxItemsPerPage&quot;&#58;8,&quot;hideWebPartWhenEmpty&quot;&#58;false,&quot;sites&quot;&#58;[],&quot;layoutId&quot;&#58;&quot;Card&quot;,&quot;dataProviderId&quot;&#58;&quot;Search&quot;,&quot;webId&quot;&#58;&quot;09f67e23-4ebd-4a02-a7f3-c0ce35767d71&quot;,&quot;siteId&quot;&#58;&quot;671077ce-0398-45bf-9f57-9aa37d0dc972&quot;&#125;&#125;\"><div data-sp-componentid=\"\">daf0b71c-6de8-4ef7-b511-faae7c388708</div><div data-sp-htmlproperties=\"\"><div data-sp-prop-name=\"title\" data-sp-searchableplaintext=\"true\">Most recent documents</div><a data-sp-prop-name=\"baseUrl\" href=\"/sites/mtest_001\"></a></div></div></div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.0\" data-sp-controldata=\"&#123;&quot;controlType&quot;&#58;4,&quot;displayMode&quot;&#58;2,&quot;id&quot;&#58;&quot;330adeae-1e33-413c-9d54-eaffbf90896a&quot;,&quot;position&quot;&#58;&#123;&quot;controlIndex&quot;&#58;1,&quot;sectionIndex&quot;&#58;2,&quot;sectionFactor&quot;&#58;4,&quot;zoneIndex&quot;&#58;1&#125;,&quot;innerHTML&quot;&#58;&quot;&lt;p&gt;Hello world&lt;/p&gt;\\n&quot;,&quot;editorType&quot;&#58;&quot;CKEditor&quot;,&quot;addedFromPersistedData&quot;&#58;true&#125;\"><div data-sp-rte=\"\"><p>Hello world</p>\n</div></div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.0\" data-sp-controldata=\"&#123;&quot;controlType&quot;&#58;3,&quot;displayMode&quot;&#58;2,&quot;id&quot;&#58;&quot;ace941e2-09d1-4acb-a1c1-dedf79ca315b&quot;,&quot;position&quot;&#58;&#123;&quot;zoneIndex&quot;&#58;1,&quot;sectionIndex&quot;&#58;2,&quot;controlIndex&quot;&#58;2,&quot;sectionFactor&quot;&#58;4&#125;,&quot;webPartId&quot;&#58;&quot;934d533d-8386-45d1-96be-3e85d16a4c83&quot;&#125;\"><div data-sp-webpart=\"\" data-sp-webpartdataversion=\"1.0\" data-sp-webpartdata=\"&#123;&quot;id&quot;&#58;&quot;934d533d-8386-45d1-96be-3e85d16a4c83&quot;,&quot;instanceId&quot;&#58;&quot;ace941e2-09d1-4acb-a1c1-dedf79ca315b&quot;,&quot;title&quot;&#58;&quot;Orders&quot;,&quot;description&quot;&#58;&quot;Shows recent orders&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.0&quot;,&quot;properties&quot;&#58;&#123;&quot;description&quot;&#58;&quot;Orders&quot;&#125;&#125;\"><div data-sp-componentid=\"\">934d533d-8386-45d1-96be-3e85d16a4c83</div><div data-sp-htmlproperties=\"\"></div></div></div></div>";

const customTeamSiteHomePage = '<div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&quot;&#58;3,&quot;displayMode&quot;&#58;2,&quot;id&quot;&#58;&quot;ede2ee65-157d-4523-b4ed-87b9b64374a6&quot;,&quot;position&quot;&#58;&#123;&quot;zoneIndex&quot;&#58;1,&quot;sectionIndex&quot;&#58;1,&quot;controlIndex&quot;&#58;0.5,&quot;sectionFactor&quot;&#58;8&#125;,&quot;webPartId&quot;&#58;&quot;34b617b3-5f5d-4682-98ed-fc6908dc0f4c&quot;&#125;"><div data-sp-webpart="" data-sp-webpartdataversion="1.0" data-sp-webpartdata="&#123;&quot;id&quot;&#58;&quot;34b617b3-5f5d-4682-98ed-fc6908dc0f4c&quot;,&quot;instanceId&quot;&#58;&quot;ede2ee65-157d-4523-b4ed-87b9b64374a6&quot;,&quot;title&quot;&#58;&quot;Minified HelloWorld&quot;,&quot;description&quot;&#58;&quot;HelloWorld description&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.0&quot;,&quot;properties&quot;&#58;&#123;&quot;description&quot;&#58;&quot;HelloWorld&quot;&#125;&#125;"><div data-sp-componentid="">34b617b3-5f5d-4682-98ed-fc6908dc0f4c</div><div data-sp-htmlproperties=""></div></div></div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&quot;&#58;3,&quot;webPartId&quot;&#58;&quot;8c88f208-6c77-4bdb-86a0-0c47b4316588&quot;,&quot;position&quot;&#58;&#123;&quot;zoneIndex&quot;&#58;1,&quot;sectionIndex&quot;&#58;1,&quot;controlIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;8&#125;,&quot;displayMode&quot;&#58;2,&quot;addedFromPersistedData&quot;&#58;true,&quot;id&quot;&#58;&quot;3ede60d3-dc2c-438b-b5bf-cc40bb2351e5&quot;&#125;"><div data-sp-webpart="" data-sp-webpartdataversion="1.0" data-sp-webpartdata="&#123;&quot;id&quot;&#58;&quot;8c88f208-6c77-4bdb-86a0-0c47b4316588&quot;,&quot;instanceId&quot;&#58;&quot;3ede60d3-dc2c-438b-b5bf-cc40bb2351e5&quot;,&quot;title&quot;&#58;&quot;News&quot;,&quot;description&quot;&#58;&quot;Display recent news.&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&quot;title&quot;&#58;&quot;News&quot;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&quot;baseUrl&quot;&#58;&quot;https&#58;//m365x656971.sharepoint.com/sites/marketing&quot;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.0&quot;,&quot;properties&quot;&#58;&#123;&quot;layoutId&quot;&#58;&quot;FeaturedNews&quot;,&quot;dataProviderId&quot;&#58;&quot;viewCounts&quot;,&quot;emptyStateHelpItemsCount&quot;&#58;1,&quot;newsDataSourceProp&quot;&#58;2,&quot;newsSiteList&quot;&#58;[],&quot;webId&quot;&#58;&quot;4f118c69-66e0-497c-96ff-d7855ce0713d&quot;,&quot;siteId&quot;&#58;&quot;016bd1f4-ea50-46a4-809b-e97efb96399c&quot;&#125;&#125;"><div data-sp-componentid="">8c88f208-6c77-4bdb-86a0-0c47b4316588</div><div data-sp-htmlproperties=""><div data-sp-prop-name="title" data-sp-searchableplaintext="true">News</div><a data-sp-prop-name="baseUrl" href="/sites/marketing"></a></div></div></div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&quot;&#58;3,&quot;webPartId&quot;&#58;&quot;c70391ea-0b10-4ee9-b2b4-006d3fcad0cd&quot;,&quot;position&quot;&#58;&#123;&quot;zoneIndex&quot;&#58;1,&quot;sectionIndex&quot;&#58;2,&quot;controlIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;4&#125;,&quot;displayMode&quot;&#58;2,&quot;addedFromPersistedData&quot;&#58;true,&quot;id&quot;&#58;&quot;63da0d97-9db4-4847-a4bf-3ae019d4c6f2&quot;&#125;"><div data-sp-webpart="" data-sp-webpartdataversion="1.0" data-sp-webpartdata="&#123;&quot;id&quot;&#58;&quot;c70391ea-0b10-4ee9-b2b4-006d3fcad0cd&quot;,&quot;instanceId&quot;&#58;&quot;63da0d97-9db4-4847-a4bf-3ae019d4c6f2&quot;,&quot;title&quot;&#58;&quot;Quick links&quot;,&quot;description&quot;&#58;&quot;Add links to important documents and pages.&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&quot;title&quot;&#58;&quot;Quick links&quot;,&quot;items[0].title&quot;&#58;&quot;Learn about a team site&quot;,&quot;items[1].title&quot;&#58;&quot;Learn how to add a page&quot;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&quot;baseUrl&quot;&#58;&quot;https&#58;//m365x656971.sharepoint.com/sites/marketing&quot;,&quot;items[0].url&quot;&#58;&quot;https&#58;//go.microsoft.com/fwlink/p/?linkid=827918&quot;,&quot;items[1].url&quot;&#58;&quot;https&#58;//go.microsoft.com/fwlink/p/?linkid=827919&quot;,&quot;items[0].renderInfo.linkUrl&quot;&#58;&quot;https&#58;//go.microsoft.com/fwlink/p/?linkid=827918&quot;,&quot;items[1].renderInfo.linkUrl&quot;&#58;&quot;https&#58;//go.microsoft.com/fwlink/p/?linkid=827919&quot;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.0&quot;,&quot;properties&quot;&#58;&#123;&quot;items&quot;&#58;[&#123;&quot;siteId&quot;&#58;&quot;00000000-0000-0000-0000-000000000000&quot;,&quot;webId&quot;&#58;&quot;00000000-0000-0000-0000-000000000000&quot;,&quot;uniqueId&quot;&#58;&quot;00000000-0000-0000-0000-000000000000&quot;,&quot;itemType&quot;&#58;2,&quot;fileExtension&quot;&#58;&quot;com/fwlink/p/?linkid=827918&quot;,&quot;progId&quot;&#58;&quot;&quot;,&quot;flags&quot;&#58;0,&quot;hasInvalidUrl&quot;&#58;false,&quot;renderInfo&quot;&#58;&#123;&quot;imageUrl&quot;&#58;&quot;&quot;,&quot;compactImageInfo&quot;&#58;&#123;&quot;iconName&quot;&#58;&quot;Globe&quot;,&quot;color&quot;&#58;&quot;&quot;,&quot;imageUrl&quot;&#58;&quot;&quot;,&quot;forceIconSize&quot;&#58;true&#125;,&quot;backupImageUrl&quot;&#58;&quot;&quot;,&quot;iconUrl&quot;&#58;&quot;&quot;,&quot;accentColor&quot;&#58;&quot;&quot;,&quot;imageFit&quot;&#58;0,&quot;forceStandardImageSize&quot;&#58;false,&quot;isFetching&quot;&#58;false&#125;,&quot;id&quot;&#58;1&#125;,&#123;&quot;siteId&quot;&#58;&quot;00000000-0000-0000-0000-000000000000&quot;,&quot;webId&quot;&#58;&quot;00000000-0000-0000-0000-000000000000&quot;,&quot;uniqueId&quot;&#58;&quot;00000000-0000-0000-0000-000000000000&quot;,&quot;itemType&quot;&#58;2,&quot;fileExtension&quot;&#58;&quot;com/fwlink/p/?linkid=827919&quot;,&quot;progId&quot;&#58;&quot;&quot;,&quot;flags&quot;&#58;0,&quot;hasInvalidUrl&quot;&#58;false,&quot;renderInfo&quot;&#58;&#123;&quot;imageUrl&quot;&#58;&quot;&quot;,&quot;compactImageInfo&quot;&#58;&#123;&quot;iconName&quot;&#58;&quot;Globe&quot;,&quot;color&quot;&#58;&quot;&quot;,&quot;imageUrl&quot;&#58;&quot;&quot;,&quot;forceIconSize&quot;&#58;true&#125;,&quot;backupImageUrl&quot;&#58;&quot;&quot;,&quot;iconUrl&quot;&#58;&quot;&quot;,&quot;accentColor&quot;&#58;&quot;&quot;,&quot;imageFit&quot;&#58;0,&quot;forceStandardImageSize&quot;&#58;false,&quot;isFetching&quot;&#58;false&#125;,&quot;id&quot;&#58;2&#125;],&quot;isMigrated&quot;&#58;true,&quot;layoutId&quot;&#58;&quot;CompactCard&quot;,&quot;shouldShowThumbnail&quot;&#58;true,&quot;hideWebPartWhenEmpty&quot;&#58;true,&quot;dataProviderId&quot;&#58;&quot;QuickLinks&quot;,&quot;webId&quot;&#58;&quot;4f118c69-66e0-497c-96ff-d7855ce0713d&quot;,&quot;siteId&quot;&#58;&quot;016bd1f4-ea50-46a4-809b-e97efb96399c&quot;&#125;&#125;"><div data-sp-componentid="">c70391ea-0b10-4ee9-b2b4-006d3fcad0cd</div><div data-sp-htmlproperties=""><div data-sp-prop-name="title" data-sp-searchableplaintext="true">Quick links</div><div data-sp-prop-name="items[0].title" data-sp-searchableplaintext="true">Learn about a team site</div><div data-sp-prop-name="items[1].title" data-sp-searchableplaintext="true">Learn how to add a page</div><a data-sp-prop-name="baseUrl" href="/sites/marketing"></a><a data-sp-prop-name="items[0].url" href="https&#58;//go.microsoft.com/fwlink/p/?linkid=827918"></a><a data-sp-prop-name="items[1].url" href="https&#58;//go.microsoft.com/fwlink/p/?linkid=827919"></a><a data-sp-prop-name="items[0].renderInfo.linkUrl" href="https&#58;//go.microsoft.com/fwlink/p/?linkid=827918"></a><a data-sp-prop-name="items[1].renderInfo.linkUrl" href="https&#58;//go.microsoft.com/fwlink/p/?linkid=827919"></a></div></div></div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&quot;&#58;3,&quot;webPartId&quot;&#58;&quot;eb95c819-ab8f-4689-bd03-0c2d65d47b1f&quot;,&quot;position&quot;&#58;&#123;&quot;zoneIndex&quot;&#58;2,&quot;sectionIndex&quot;&#58;1,&quot;controlIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;8&#125;,&quot;displayMode&quot;&#58;2,&quot;addedFromPersistedData&quot;&#58;true,&quot;id&quot;&#58;&quot;4366ceff-b92b-4a12-905e-1dd2535f976d&quot;&#125;"><div data-sp-webpart="" data-sp-webpartdataversion="1.0" data-sp-webpartdata="&#123;&quot;id&quot;&#58;&quot;eb95c819-ab8f-4689-bd03-0c2d65d47b1f&quot;,&quot;instanceId&quot;&#58;&quot;4366ceff-b92b-4a12-905e-1dd2535f976d&quot;,&quot;title&quot;&#58;&quot;Site activity&quot;,&quot;description&quot;&#58;&quot;Show recent activities from your site.&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.0&quot;,&quot;properties&quot;&#58;&#123;&quot;maxItems&quot;&#58;9&#125;&#125;"><div data-sp-componentid="">eb95c819-ab8f-4689-bd03-0c2d65d47b1f</div><div data-sp-htmlproperties=""></div></div></div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&quot;&#58;3,&quot;webPartId&quot;&#58;&quot;f92bf067-bc19-489e-a556-7fe95f508720&quot;,&quot;position&quot;&#58;&#123;&quot;zoneIndex&quot;&#58;2,&quot;sectionIndex&quot;&#58;2,&quot;controlIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;4&#125;,&quot;addedFromPersistedData&quot;&#58;true,&quot;displayMode&quot;&#58;2,&quot;id&quot;&#58;&quot;456dfbc7-57be-4489-92ce-666224c4fcf1&quot;&#125;"><div data-sp-webpart="" data-sp-webpartdataversion="1.0" data-sp-webpartdata="&#123;&quot;id&quot;&#58;&quot;f92bf067-bc19-489e-a556-7fe95f508720&quot;,&quot;instanceId&quot;&#58;&quot;456dfbc7-57be-4489-92ce-666224c4fcf1&quot;,&quot;title&quot;&#58;&quot;Document library&quot;,&quot;description&quot;&#58;&quot;Add a document library.&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.0&quot;,&quot;properties&quot;&#58;&#123;&quot;isDocumentLibrary&quot;&#58;true,&quot;showDefaultDocumentLibrary&quot;&#58;true,&quot;webpartHeightKey&quot;&#58;4,&quot;selectedListUrl&quot;&#58;&quot;&quot;,&quot;listTitle&quot;&#58;&quot;Documents&quot;&#125;&#125;"><div data-sp-componentid="">f92bf067-bc19-489e-a556-7fe95f508720</div><div data-sp-htmlproperties=""></div></div></div></div>';

const layoutWebpartsContent = "<div><div data-sp-canvascontrol=\"\" data-sp-canvasdataversion=\"1.3\" data-sp-controldata=\"&#123;&quot;id&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;instanceId&quot;&#58;&quot;cbe7b0a9-3504-44dd-a3a3-0e5cacd07788&quot;,&quot;title&quot;&#58;&quot;Title Region&quot;,&quot;description&quot;&#58;&quot;Title Region Description&quot;,&quot;serverProcessedContent&quot;&#58;&#123;&quot;htmlStrings&quot;&#58;&#123;&#125;,&quot;searchablePlainTexts&quot;&#58;&#123;&#125;,&quot;imageSources&quot;&#58;&#123;&#125;,&quot;links&quot;&#58;&#123;&#125;&#125;,&quot;dataVersion&quot;&#58;&quot;1.3&quot;,&quot;properties&quot;&#58;&#123;&quot;title&quot;&#58;&quot;Page 1&quot;,&quot;textAlignCenter&quot;&#58;false,&quot;imageSourceType&quot;&#58;4,&quot;translateX&quot;&#58;50,&quot;translateY&quot;&#58;50&#125;&#125;\"></div></div>";

const documentFragment = parse5.parseFragment(customTeamSiteHomePage);
const controlElements = [].concat(getControlElements(documentFragment));
const controls = getControls(controlElements);
console.log(JSON.stringify(controls, null, 2));

function getControls(controlElements) {
  const controls = [];
  let controlOrder = 0;
  for (let i = 0; i < controlElements.length; i++) {
    const controlDataString = getAttribute(controlElements[i], 'data-sp-controldata');
    if (!controlDataString) {
      continue;
    }

    const controlData = JSON.parse(controlDataString);
    const control = {
      order: controlOrder,
      controlType: getControlType(controlData.controlType),
      instanceId: controlData.id,
      dataVersion: getAttribute(controlElements[i], 'data-sp-canvasdataversion'),
      canvasControlData: getAttribute(controlElements[i], 'data-sp-canvascontrol'),
      spControlData: controlData
    }
    switch (controlData.controlType) {
      // clientSideWebPart
      case 3:
        const wpDiv = getWpDiv(controlElements[i]);
        if (wpDiv) {
          const wpDataRaw = getAttribute(wpDiv, 'data-sp-webpartdata');
          if (wpDataRaw) {
            control.webPartData = wpDataRaw;
            const wpData = JSON.parse(wpDataRaw);
            control.title = wpData.title;
            control.description = wpData.description;
            control.properties = wpData.properties;
            control.serverProcessedContent = wpData.serverProcessedContent;
            control.webPartId = wpData.id;
          }

          const htmlPropertiesDiv = getHtmlPropertiesDiv(controlElements[i]);
          if (htmlPropertiesDiv) {
            control.htmlPropertiesData = parse5.serialize(htmlPropertiesDiv);
            control.htmlProperties = getAttribute(htmlPropertiesDiv, 'data-sp-htmlproperties');
          }
        }
        break;
      // clientSideText
      case 4:
        // todo: https://github.com/SharePoint/PnP-Sites-Core/blob/0dfd271515f0e5cac24e7165416ecf2320e07081/Core/OfficeDevPnP.Core/Pages/ClientSidePageControls.cs#L488
        break;
      // canvasColumn
      case 0:
        // no additional processing needed
        break;
    }
    controls.push(control);
    controlOrder++;
  }

  return controls;
}

function getWpDiv(elem) {
  if (elem.nodeName === 'div' &&
    typeof getAttribute(elem, 'data-sp-webpartdata') !== 'undefined') {
    return elem;
  }

  if (!elem.childNodes) {
    return undefined;
  }

  for (let i = 0; i < elem.childNodes.length; i++) {
    const node = elem.childNodes[i];
    const wpDiv = getWpDiv(node);
    if (wpDiv) {
      return wpDiv;
    }
  }
}

function getHtmlPropertiesDiv(elem) {
  if (elem.nodeName === 'div' &&
    typeof getAttribute(elem, 'data-sp-htmlproperties') !== 'undefined') {
    return elem;
  }

  if (!elem.childNodes) {
    return undefined;
  }

  for (let i = 0; i < elem.childNodes.length; i++) {
    const node = elem.childNodes[i];
    const wpDiv = getHtmlPropertiesDiv(node);
    if (wpDiv) {
      return wpDiv;
    }
  }
}

function getControlElements(elem) {
  let controlElements = [];
  let childNodes = elem.childNodes;
  if (!childNodes) {
    return controlElements;
  }

  for (let i = 0; i < childNodes.length; i++) {
    const node = childNodes[i];
    const controlData = getAttribute(node, 'data-sp-controldata');
    if (typeof controlData !== 'undefined') {
      controlElements.push(node);
    }
    else {
      controlElements = controlElements.concat(getControlElements(node));
    }
  }

  return controlElements;
}

function getAttribute(elem, name) {
  const attrs = elem.attrs;

  if (!attrs) {
    return undefined;
  }

  for (let i = 0; i < attrs.length; i++) {
    if (attrs[i].name === name) {
      return attrs[i].value;
    }
  }

  return undefined;
}

function getControlType(controlTypeNumber) {
  switch (controlTypeNumber) {
    case 0:
      return 'CanvasColumn';
    case 3:
      return 'ClientSideWebPart';
    case 4:
      return 'ClientSideText';
    default:
      return null;
  }
}