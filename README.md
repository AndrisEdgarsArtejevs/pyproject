# pyproject

Mērķis: 
    Uzrādīt viegli pāredzamā veidā izvēlētā studenta vidējos semestra vērtējumus. Atzīmju analīze tiks uzrādīta diagrammā izmantojot aplikāciju Excel.

Darba gaita:
    Sākotnējā ideja bija nolasīt datus no e-klases izmantojot metodi scraper, bet diemžēl tas nebija iespējams, jo vietnei ir uzlikts blocks(drošība). Tākā scrapers nesanāca izdomāju lasīt datus izmantjot pdf failu, ko es arī šajā programā īstenoju. Vēlējos nolasīt vairāku studentu datus un uztaisīt milzīgu masīvu kurā tos es samestu, un tad izmantojot visus datus, es izveidotu atzīmju analīzi programmā Excel. Diemžēl nonācu pie dažiem klupšanas akmeņiem un pašreizējā programma spēj vienīgi nolasīt datus vienam studentam. Programma nolasa datus no izvēlētā pdf faila, iztīra tos un ievieto masīvā. Masīvs tiek apstrādāts tā, lai dati būtu lietojami exel failā, ar to es domāju, lai no datiem būtu iespēja izveidot diagramu kurā tiek attēlotas vidējās atzīmes semestrī. 

Imports:
    import PyPDF2 - Tiek lietots, lai izvilktu ārā datus no pdf faila
    from openpyxl import Workbook - Tiek lietots, lai pikļūtu un manipulētu ar excel datiem
    from openpyxl.chart import BarChart, Reference - Tiek lietots, lai izveidotu excel diagrammas
    from openpyxl.chart.label import DataLabelList - Tiek lietots, lai piesšķirtu etiķetes diagrammām

references(sources):
    https://www.geeksforgeeks.org/python-plotting-charts-in-excel-sheet-using-openpyxl-module-set-3/
    https://www.youtube.com/watch?v=vsrxkJ9HF24
    https://www.w3schools.com/python/default.asp