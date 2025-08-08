
# Pub Crawl Route Optimizer (Excel VBA, Dijkstra)

Excel-based Decision Support System (DSS) that computes the **shortest path** between any two pubs in Bath, showing intermediate stops and cumulative distances. Built for the *Optimisation & Spreadsheet Modelling* module (University of Bath).

## ✨ Features
- Shortest path using **Dijkstra’s Algorithm**
- **Geodesic** distance calculation (lat/long)
- Clean UI: dropdown **Origin/Destination**, **Start** & **Reset** buttons
- Output table: Pub Name, Pub ID, Distance from Origin, Link ID
- Error-handling & data validation

## 🛠 Tech
**Excel (.xlsm), VBA**, Geodesic distance function, Dijkstra’s algorithm
[Report.pdf](https://github.com/user-attachments/files/21676743/Report.pdf)
[Report.pdf](https://github.com/user-attachments/files/21676735/Report.pdf)
[MN50750 Assignment 2 - Building a Decision Support System For Pub Crawlers.pdf](https://github.com/user-attachments/files/21676733/MN50750.Assignment.2.-.Building.a.Decision.Support.System.For.Pub.Crawlers.pdf)

## 📦 Files
- `PubCrawl_DSS.xlsm` – main tool (enable macros)
- `Report.pdf` – project report / method
- `media/` – screenshots and optional demo gif

## ▶️ How to Run
1. Download `PubCrawl_DSS.xlsm`
2. Open in **Excel (Windows)** and **Enable Macros**
3. On the *Optimal Path* sheet:<img width="460" height="159" alt="Screenshot 2025-08-08 040050" src="https://github.com/user-attachments/assets/f9a61bae-a5ab-487e-bb11-8ff38dc4f124" />

   - Pick **Origin Pub** and **Destination Pub** from dropdowns
   - Click **Start** to compute shortest route
   - Use **Reset** to clear and run another trial

> Tested in Microsoft Excel (Windows). Mac Excel may handle VBA differently.

## 📚 Method
- Graph nodes = pubs; edges = roads with geodesic distances  
- Dijkstra’s algorithm to compute shortest path  
- Output includes cumulative distance per stop + Link IDs

## 🚧 Limitations / Next Steps
- Static data (no live closures/traffic)
- Bath-only dataset (could generalize with new city sheets)
- Add map visual / web or mobile front end

## 📝 License
MIT © Umar Farooq
