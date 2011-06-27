function CFileName(strFileName){
		var arrPath = strFileName.split('\\');
		var strLast = arrPath[arrPath.length - 1];
		var arrFile = strLast.split('.');
		var strName = '';
		for (var i=0; i<arrFile.length - 1; i++)
		{
			if (i > 0)
			{
				strName += '.';
			}
			strName += arrFile[i];
		}
		
		var strExt = '';
		if (arrFile.length > 1)
		{
			strExt = arrFile[arrFile.length - 1];
		}

		return {
			FileName: strName,
			Extend: strExt
		};
}