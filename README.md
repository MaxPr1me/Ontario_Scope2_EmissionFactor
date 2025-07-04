# Ontario Scope2 Emission Factor

This repository processes IESO data to compute electricity emission factors for Ontario. It originally shipped as a Google Colab notebook but now also includes a standalone Python script that can be run locally or via GitHub Actions.

## Running locally
1. Place the required IESO data files inside `data/IESO_Data` using the same folder structure expected by the notebook (e.g. `Supply/GOC-2016.xlsx`).
2. Install the dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Execute one of the scripts:
   ```bash
   # Original emission factor workflow
   python src/emission_factors.py

   # Converted analysis from Scott_Code_Feb2025.ipynb
   python src/scott_code.py
   ```

A simple GitHub Actions workflow is provided to demonstrate running the script automatically on each push.

## Citation

If you use this code in your research or publication, please cite the following paper:

> St-Jacques, M., Bucking, S., & O'Brien, W. (2020). Spatially and temporally sensitive consumption-based emission factors from mixed-use electrical grids for building electrical use. *Energy and Buildings*, 224, 110249. https://doi.org/10.1016/j.enbuild.2020.110249

**BibTeX:**
```bibtex
@article{STJACQUES2020110249,
  title = {Spatially and temporally sensitive consumption-based emission factors from mixed-use electrical grids for building electrical use},
  author = {Max St-Jacques and Scott Bucking and William O'Brien},
  journal = {Energy and Buildings},
  volume = {224},
  pages = {110249},
  year = {2020},
  issn = {0378-7788},
  doi = {https://doi.org/10.1016/j.enbuild.2020.110249},
  url = {https://www.sciencedirect.com/science/article/pii/S0378778819337387}
}
```
