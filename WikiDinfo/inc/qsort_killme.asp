<%
option explicit
const CMP_LESS = -1
const CMP_EQU = 0
const CMP_GREATER = 1

function qsort_cmp(a,b)
	if a < b then 
		qsort_cmp = CMP_LESS
	elseif a > b then 
		qsort_cmp = CMP_GREATER
	else 
		qsort_cmp = CMP_EQU
	end if
end function

function array_swap(values, i,j)
	dim temp: temp = values(i)
	values(i) = values (j)
	values(j) = temp
end function

function array_swapO(values, i,j)
	dim temp : set temp = values(i)
	set values(i) = values (j)
	set values(j) = temp
end function

class QSort
	private f_cmp
	private f_swap
	private valueSort
	
	private sub class_initialize
	end sub
	
	public Sub QSort(values, loBound,hiBound)
	' This function derived from: 
	' 	http://4guysfromrolla.com/webtech/012799-2.shtml
		Dim pivot,loSwap,hiSwap

	  '== Two items to sort
		if hiBound - loBound = 1 then
			if f_cmp(values(loBound), values(hiBound)) = CMP_GREATER then
				f_swap values,loBound,hiBound
			End if
			exit sub
		End If
	
	  '== Three or more items to sort
		dim pivotIndex : pivotIndex = int((loBound + hiBound) / 2)
		
		if valueSort then
			pivot = values(pivotIndex)
		else
			set pivot = values(pivotIndex)
		end if
		
		f_swap values, pivotIndex, loBound
		
		loSwap = loBound + 1
		hiSwap = hiBound
	  
		do
			'== Find the right loSwap
			while (loSwap < hiSwap) and (f_cmp(values(loSwap), pivot) <> CMP_GREATER)
				loSwap = loSwap + 1
			wend
			'== Find the right hiSwap
			while (f_cmp(values(hiSwap), pivot) = CMP_GREATER) 'values(hiSwap) > pivot
				hiSwap = hiSwap - 1
			wend
			'== Swap values if loSwap is less then hiSwap
			if loSwap < hiSwap then
				f_swap values, loSwap, hiSwap
			End If
		loop while loSwap < hiSwap
	  
		if valueSort then
			values(loBound) = values(hiSwap)
			values(hiSwap) = pivot
		else
			set values(loBound) = values(hiSwap)
			set values(hiSwap) = pivot
		end if
	  
		'== Recursively call function .. the beauty of Quicksort
		'== 2 or more items in first section
		if loBound < (hiSwap - 1) then QSort values, loBound,hiSwap-1
		'== 2 or more items in second section
		if hiSwap + 1 < hibound then QSort values, hiSwap+1,hiBound
	End Sub

	public property set Compare(func)
		set f_cmp = func
	end property
	
	public property get Compare
		set compare = f_cmp
	end property
	
	public sub Sort(values)
		' Don't sort empty arrays or arrays with only 1 value
		if UBound(values) < 1 then exit sub
		valueSort = false
		
		if IsEmpty(f_cmp) then
			valueSort = true
			set f_cmp = GetRef("qsort_cmp")
			set f_swap = GetRef("array_swap")
		else
			set f_swap = GetRef("array_swapO")
		end if
		
		QSort values, LBound(values), UBound(values)

		if valueSort then f_cmp = Empty
	end sub
end class

dim a : a = Array(4,1,9,2,7)
dim sort : set sort = new QSort
sort.Sort a

dim i
for i = 0 to UBound(a)
 if (i > 0) then Response.Write ","
 Response.Write a(i)
next
%>

