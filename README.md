# Ontario Scope2 Emission Factor

This repository processes IESO data to compute electricity emission factors for Ontario using a standalone Python script.

## Running locally
1. Install the dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Ensure the files `data/emission_rates.csv` and `data/neighboring_emission_factors.csv` exist. Edit them if you need to change the default values.
   The script downloads the required IESO data automatically.
3. Run the emission factor script:
   ```bash
   python src/Ontario_EF_Code.py
   ```

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
