This Excel Add-in is an implementation of the classic Metric MDS model.

This model places points into a 2D or 3D map, based on pairwise distances among these points.  This tool is commonly used for perceptual mapping, where the marketer wants to display brands on a map based on consumer perceptions of similarities among these brands.  Please keep in mind that the model takes the data as pairwise distances, so that larger distances will result into the respective points being farther away. If you have similarity, rather than distance data, you will need to transform them (e.g., invert their signs).

The files needed to run the model are:

Metric MDS 2003.xla or Metric MDS 2007.xlam - Excel Add-in
MDSDLL.dll - dynamic link library that does all the heavy-dutty number crunching.

HOW TO INSTALL:

1. Copy the Metric MDS.xla (or Metric MDS 2007.xlam if you are using Office 2007) file into its own directory
2. Copy the MDSDLL.dll file into your Windows/System32 (or Windows/System for Windows 7) directory
3. Make sure that you already have the "Analysis Toolpack" and "Analysis Toolpack VBA" Excel add-ins installed. 
4. Open any Excel sheet (e.g., the attached Banknotes.xls sheet)
5. Check the Tools/Add-ins/Browse option and point it to the MetricMDS.xla (or xlam for Office 2007) file.

If you want to demonstrate how metric MDS works, you may also find the "Magazine Similarity" worksheet useful.  It gathers data from one respondent and then produces a perceptual map from these data using the Metric MDS model.

For details about Multidimensional Scaling:

Kruskal, J. B., and Wish, M. (1978), Multidimensional Scaling, Sage University Paper series on Quantitative Application in the Social Sciences, 07-011. Beverly Hills and London: Sage Publications. 