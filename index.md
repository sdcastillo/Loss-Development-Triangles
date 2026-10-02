---
layout: default
title: Loss Development Triangles
description: Merge claims listings into a cumulative paid triangle with R.
samwiki: true
---

<section class="sw-lede" aria-labelledby="triangle-title">
  <div class="sw-lede-copy">
    <h2 id="triangle-title">Merging claims listings into a loss development triangle</h2>
    <p>A loss development triangle is the schedule actuaries use to see how losses from a fixed accident period mature as later valuations arrive. Rows are origin periods. Columns are development ages. The cell <em>C</em><sub><em>i</em>,<em>k</em></sub> is cumulative paid loss for accident year <em>i</em> observed at age <em>k</em>. Cells already valued form the upper triangle. Reserve methods fill the lower triangle by carrying each origin forward with age-to-age factors. The <a href="https://en.wikipedia.org/wiki/Chain-ladder_method">chain-ladder method</a> is the usual reading of that array.</p>
    <p>The inputs are claims listings, one Excel workbook per evaluation year. Each row is a claim, with <code>accident_year</code>, paid-to-date in <code>paid</code>, and <code>file_year</code> for the year of that valuation. Listings are not required to share unused columns. A later file can carry extra fields, and the merge keeps only the three columns that define the triangle. Stacking those snapshots in R replaces a workbook of cross-file links that has to be rewired at every new evaluation.</p>
    <p>Development age is the lag from the accident year to the valuation year. For annual valuations the lag in months is (<code>file_year</code> − <code>accident_year</code>) × 12, so a 2011 accident valued in 2011 sits at age 0 and the same accident valued in 2012 sits at age 12. Paid in each listing is paid-to-date at that <code>file_year</code>, not the incremental payment since the previous file. Summing <code>paid</code> over claims that share an accident year and a development age produces the cumulative cell <em>C</em><sub><em>i</em>,<em>k</em></sub>. The same origin then occupies one column per valuation instead of being overwritten.</p>
    <p>Age-to-age factors compare cumulative paid at successive ages that appear in the triangle. The development ages are the month lags themselves, not a unit index, so the columns of this construction are 0, 12, 24, and so on when valuations are annual. Write the observed ages as d₁ &lt; d₂ &lt; … &lt; dₘ. For each consecutive pair, let Sⱼ be the accident years for which both Cᵢ,dⱼ and Cᵢ,dⱼ₊₁ are present. The volume-weighted link ratio is fⱼ = (Σᵢ∈Sⱼ Cᵢ,dⱼ₊₁) / (Σᵢ∈Sⱼ Cᵢ,dⱼ). The cumulative development factor from age dⱼ to the last observed age is the product fⱼ fⱼ₊₁ … fₘ₋₁. An origin missing an intermediate valuation stays out of any factor that needs the missing cell.</p>
    <p>The script selects <code>file_year</code>, <code>accident_year</code>, and <code>paid</code> from every workbook with <code>readxl</code>, stacks the frames with <code>ldply</code>, and attaches <code>maturity_in_months</code>. <code>ChainLadder::as.triangle</code> is called with <code>origin = "accident_year"</code>, <code>dev = "maturity_in_months"</code>, and <code>value = "paid"</code>. Its default aggregation sums <code>paid</code> inside each origin-age cell, which is the map from a claim listing to <em>C</em><sub><em>i</em>,<em>k</em></sub>. Link ratios and ultimates are a later call to <code>ata</code> and <code>cdf</code> on that triangle. The workbooks and <code>Final_Article.Rmd</code> are in this repository; a longer walkthrough is on <a href="https://datascienceplus.com/faster-than-excel-painlessly-merge-data-into-actuarial-loss-development-triangles-with-r/">DataScience+</a>.</p>
  </div>
  <aside class="sw-find" aria-labelledby="triangle-facts">
    <h2 id="triangle-facts">Triangle notation</h2>
    <ul>
      <li><strong>Origin i</strong> Accident year of the claim, stored as <code>accident_year</code>.</li>
      <li><strong>Age in months</strong> (file_year − accident_year) × 12. Annual valuations land on 0, 12, 24, …</li>
      <li><strong>Cumulative paid Cᵢ,ₖ</strong> Sum of <code>paid</code> over claims in accident year i valued at age k months.</li>
      <li><strong>Link ratio fⱼ</strong> (Σᵢ∈Sⱼ Cᵢ,dⱼ₊₁) / (Σᵢ∈Sⱼ Cᵢ,dⱼ), over origins observed at both of a consecutive pair of ages.</li>
      <li><strong>Development factor</strong> Product of the link ratios from the current age through the last age in the triangle.</li>
    </ul>
  </aside>
</section>
