	
Statistical Distributions
By
Bill Young



Contents
Tests	4
Shapiro-Wilk Normality Test	4
Kolmogorov-Smirnov	4
General Information	4
Discrete Distributions	4
Definition:	4
With Finite Support	4
With Infinite Support	5
Continuous Distributions	6
Definition:	6
Supported On A Bounded Interval	6
Supported On Intervals Of Length 2π – Directional Distributions	12
Supported On Semi-infinite Intervals, Usually [0,∞)	17
Supported On The Whole Real Line	18
With Variable Support	20
Mixed Discrete/Continuous Distributions	20
Joint Distributions	20
Two or More Random Variables On The Sample Space	20
Matrix-Valued Distributions	21
Non-numeric Distributions	21
Miscellaneous Distributions	21
Works Cited	22





#Tests
##Shapiro-Wilk Normality Test
 (Zaiontz, Shapiro-Wilk Original Test) (Zaiontz, Shapiro-Wilk Expanded Test)
##Kolmogorov-Smirnov
(Sonnier)
#General Information
Common Probability Distributions (Joyce, 2016)

#Discrete Distributions
##Definition: 
“A [statistical distribution ](http://mathworld.wolfram.com/StatisticalDistribution.html)whose variables can take on only discrete values. Abramowitz and Stegun (1972, p. 929) give a table of the parameters of most common discrete distributions.” (Weisstein, Discrete Distribution, 2017)
##With Finite Support
  *The Bernoulli distribution, which takes value 1 with probability p and value 0 with probability q = 1 − p.
  * / (Weisstein, Wolfram Mathworld, 2017)
  *The Rademacher distribution, which takes value 1 with probability 1/2 and value −1 with probability 1/2.
  *The binomial distribution, which describes the number of successes in a series of independent Yes/No experiments all with the same probability of success.
  *The beta-binomial distribution, which describes the number of successes in a series of independent Yes/No experiments with heterogeneity in the success probability.
  *The degenerate distribution at x0, where X is certain to take the value x0. This does not look random, but it satisfies the definition of random variable. This is useful because it puts deterministic variables and random variables in the same formalism.
  *The discrete uniform distribution, where all elements of a finite set are equally likely. This is the theoretical distribution model for a balanced coin, an unbiased die, a casino roulette, or the first card of a well-shuffled deck.
  *The hypergeometric distribution, which describes the number of successes in the first m of a series of n consecutive Yes/No experiments, if the total number of successes is known. This distribution arises when there is no replacement.
  *The Poisson binomial distribution, which describes the number of successes in a series of independent Yes/No experiments with different success probabilities.
  *Fisher's noncentral hypergeometric distribution
  *Wallenius' noncentral hypergeometric distribution
  *Benford's law, which describes the frequency of the first digit of many naturally occurring data.

##With Infinite Support
  *The beta negative binomial distribution
  *The Boltzmann distribution, a discrete distribution important in statistical physics which describes the probabilities of the various discrete energy levels of a system in thermal equilibrium. It has a continuous analogue. Special cases include: 
  *The Gibbs distribution
  *The Maxwell–Boltzmann distribution
  *The Borel distribution
  *The Champernowne distribution
  *The extended negative binomial distribution
  *The extended hypergeometric distribution
  *The generalized log-series distribution
  *The geometric distribution, a discrete distribution which describes the number of attempts needed to get the first success in a series of independent Bernoulli trials, or alternatively only the number of losses before the first success (i.e. one less).
  *The logarithmic (series) distribution
  *The negative binomial distribution or Pascal distribution a generalization of the geometric distribution to the nth success.
  *The discrete compound Poisson distribution
  *The parabolic fractal distribution
  *The Poisson distribution, which describes a very large number of individually unlikely events that happen in a certain time interval. Related to this distribution are a number of other distributions: the displaced Poisson, the hyper-Poisson, the general Poisson binomial and the Poisson type distributions. 
  *The Conway–Maxwell–Poisson distribution, a two-parameter extension of the Poisson distribution with an adjustable rate of decay.
  *The Zero-truncated Poisson distribution, for processes in which zero counts are not observed
  *The Polya–Eggenberger distribution
  *The Skellam distribution, the distribution of the difference between two independent Poisson-distributed random variables.
  *The skew elliptical distribution
  *The Yule–Simon distribution
  *The zeta distribution has uses in applied statistics and statistical mechanics, and perhaps may be of interest to number theorists. It is the Zipf distribution for an infinite number of elements.
  *Zipf's law or the Zipf distribution. A discrete power-law distribution, the most famous example of which is the description of the frequency of words in the English language.
  *The Zipf–Mandelbrot law is a discrete power law distribution which is a generalization of the Zipf distribution.

#Continuous Distributions
##Definition:
 A continuous random variable is a random variable with a set of possible values (known as the range or support) that is infinite and uncountable. Probabilities of continuous random variables (X) are defined as the area under the curve of its distribution. Thus, only ranges of values can have a nonzero probability.  (Minitab)
##Supported On A Bounded Interval
  *The arcsine distribution on [a,b], which is a special case of the Beta distribution if a = 0 and b = 1.
  *The Beta distribution on [0,1], a family of two-parameter distributions with one mode, of which the uniform distribution is a special case, and which is useful in estimating success probabilities.
  *The logitnormal distribution on (0,1).
  */ (Wikipedia)
  *The Dirac delta function although not strictly a function, is a limiting form of many continuous probability functions. It represents a discrete probability distribution concentrated at 0 — a degenerate distribution — but the notation treats it as if it were a continuous distribution.
  *The continuous uniform distribution or rectangular distribution on [a,b], where all points in a finite interval are equally likely.
  */ (Wikipedia)
  *The Irwin–Hall distribution is the distribution of the sum of n i.i.d. U(0,1) random variables.
  */ (Wikipedia)
  *The Bates distribution is the distribution of the mean of n i.i.d. U(0,1) random variables.
  *The Kent distribution on the three-dimensional sphere.
  */ (Wikipedia)
  *The Kumaraswamy distribution is as versatile as the Beta distribution but has simple closed forms for both the cdf and the pdf.
  */ (Wikipedia)
  *The logarithmic distribution (continuous)
  *The Marchenko–Pastur distribution is important in the theory of random matrices.
  *The PERT distribution is a special case of the beta distribution
  *The raised cosine distribution
  */ (Wikipedia)
  *The reciprocal distribution
  *The triangular distribution on [a, b], a special case of which is the distribution of the sum of two independent uniformly distributed random variables (the convolution of two uniform distributions).
  */ (Wikipedia)
  *The trapezoidal distribution
  *The truncated normal distribution on [a, b].
  *The U-quadratic distribution on [a, b].
  *The von Mises-Fisher distribution on the N-dimensional sphere has the von Mises distribution as a special case.
  *The Wigner semicircle distribution is important in the theory of random matrices.

##Supported On Intervals Of Length 2π – Directional Distributions
  *The von Mises distribution
  */ (Wikipedia)
  *The wrapped normal distribution
  */ (Wikipedia)
  *The wrapped exponential distribution
  */ (Wikipedia)
  *The wrapped Lévy distribution
  *The wrapped Cauchy distribution
  */ (Wikipedia)
  *The wrapped Laplace distribution
  *The wrapped asymmetric Laplace distribution
  *
  */ (Wikipedia)
  *The Dirac comb of period 2 π although not strictly a function, is a limiting form of many directional distributions. It is essentially a wrapped Dirac delta function. It represents a discrete probability distribution concentrated at 2πn — a degenerate distribution — but the notation treats it as if it were a continuous distribution.
  */ (Wikipedia)

##Supported On Semi-infinite Intervals, Usually [0,∞)
  *The Beta prime distribution
  *The Birnbaum–Saunders distribution, also known as the fatigue life distribution, is a probability distribution used extensively in reliability applications to model failure times.
  *The chi distribution 
  *The noncentral chi distribution
  *The chi-squared distribution, which is the sum of the squares of n independent Gaussian random variables. It is a special case of the Gamma distribution, and it is used in goodness-of-fit tests in statistics. 
  *The inverse-chi-squared distribution
  *The noncentral chi-squared distribution
  *The Scaled-inverse-chi-squared distribution
  *The Dagum distribution
  *The exponential distribution, which describes the time between consecutive rare random events in a process with no memory.
  *The Exponential-logarithmic distribution
  *The F-distribution, which is the distribution of the ratio of two (normalized) chi-squared-distributed random variables, used in the analysis of variance. It is referred to as the beta prime distribution when it is the ratio of two chi-squared variates which are not normalized by dividing them by their numbers of degrees of freedom. 
  *The noncentral F-distribution
  *Fisher's z-distribution
  *The folded normal distribution
  *The Fréchet distribution
  *The Gamma distribution, which describes the time until n consecutive rare random events occur in a process with no memory. 
  *The Erlang distribution, which is a special case of the gamma distribution with integral shape parameter, developed to predict waiting times in queuing systems
  *The inverse-gamma distribution
  *The Generalized gamma distribution
  *The generalized Pareto distribution
  *The Gamma/Gompertz distribution
  *The Gompertz distribution
  *The half-normal distribution
  *Hotelling's T-squared distribution
  *The inverse Gaussian distribution, also known as the Wald distribution
  *The Lévy distribution
  *The log-Cauchy distribution
  *The log-Laplace distribution
  *The log-logistic distribution
  *The log-normal distribution, describing variables which can be modelled as the product of many small independent positive variables.
  *The Lomax distribution
  *The Mittag-Leffler distribution
  *The Nakagami distribution
  *The Pareto distribution, or "power law" distribution, used in the analysis of financial data and critical behavior.
  *The Pearson Type III distribution
  *The Phase-type distribution, used in queueing theory
  *The phased bi-exponential distribution is commonly used in pharmokinetics
  *The phased bi-Weibull distribution
  *The Rayleigh distribution
  *The Rayleigh mixture distribution
  *The Rice distribution
  *The shifted Gompertz distribution
  *The type-2 Gumbel distribution
  *The Weibull distribution or Rosin Rammler distribution, of which the exponential distribution is a special case, is used to model the lifetime of technical devices and is used to describe the particle size distribution of particles generated by grinding, milling and crushing operations.

##Supported On The Whole Real Line
  *The Behrens–Fisher distribution, which arises in the Behrens–Fisher problem.
  *The Cauchy distribution, an example of a distribution which does not have an expected value or a variance. In physics it is usually called a Lorentzian profile, and is associated with many processes, including resonance energy distribution, impact and natural spectral line broadening and quadratic stark line broadening.
  *Chernoff's distribution
  *The Exponentially modified Gaussian distribution, a convolution of a normal distribution with an exponential distribution.
  *The Fisher–Tippett, extreme value, or log-Weibull distribution
  *Fisher's z-distribution
  *The skewed generalized t distribution
  *The generalized logistic distribution
  *The generalized normal distribution
  *The geometric stable distribution
  *The Gumbel distribution
  *The Holtsmark distribution, an example of a distribution that has a finite expected value but infinite variance.
  *The hyperbolic distribution
  *The hyperbolic secant distribution
  *The Johnson SU distribution
  *The Landau distribution
  *The Laplace distribution
  *The Lévy skew alpha-stable distribution or stable distribution is a family of distributions often used to characterize financial data and critical behavior; the Cauchy distribution, Holtsmark distribution, Landau distribution, Lévy distribution and normal distribution are special cases.
  *The Linnik distribution
  *The logistic distribution
  *The map-Airy distribution
  *The normal distribution, also called the Gaussian or the bell curve. It is ubiquitous in nature and statistics due to the central limit theorem: every variable that can be modelled as a sum of many small independent, identically distributed variables with finite mean and variance is approximately normal.
  *The Normal-exponential-gamma distribution
  *The Normal-inverse Gaussian distribution
  *The Pearson Type IV distribution (see Pearson distributions)
  *The skew normal distribution
  *Student's t-distribution, useful for estimating unknown means of Gaussian populations. 
  *The noncentral t-distribution
  *The skew t distribution
  *The type-1 Gumbel distribution
  *The Tracy–Widom distribution
  *The Voigt distribution, or Voigt profile, is the convolution of a normal distribution and a Cauchy distribution. It is found in spectroscopy when spectral line profiles are broadened by a mixture of Lorentzian and Doppler broadening mechanisms.
  *The Gaussian minus exponential distribution is a convolution of a normal distribution with (minus) an exponential distribution.
  *The Chen distribution.

##With Variable Support
  *The generalized extreme value distribution has a finite upper bound or a finite lower bound depending on what range the value of one of the parameters of the distribution is in (or is supported on the whole real line for one special value of the parameter
  *The generalized Pareto distribution has a support which is either bounded below only, or bounded both above and below
  *The Tukey lambda distribution is either supported on the whole real line, or on a bounded interval, depending on what range the value of one of the parameters of the distribution is in.
  *The Wakeby distribution

#Mixed Discrete/Continuous Distributions
  *The rectified Gaussian distribution replaces negative values from a normal distribution with a discrete component at zero.
  *The compound poisson-gamma or Tweedie distribution is continuous over the strictly positive real numbers, with a mass at zero.

#Joint Distributions
##Two or More Random Variables On The Sample Space
  *The Dirichlet distribution, a generalization of the beta distribution.
  *The Ewens's sampling formula is a probability distribution on the set of all partitions of an integer n, arising in population genetics.
  *The Balding–Nichols model
  *The multinomial distribution, a generalization of the binomial distribution.
  *The multivariate normal distribution, a generalization of the normal distribution.
  *The multivariate t-distribution, a generalization of the Student's t-distribution.
  *The negative multinomial distribution, a generalization of the negative binomial distribution.
  *The generalized multivariate log-gamma distribution

##Matrix-Valued Distributions
  *The Wishart distribution
  *The inverse-Wishart distribution
  *The matrix normal distribution
  *The matrix t-distribution

#Non-numeric Distributions
#  *The categorical distribution
#Miscellaneous Distributions
  *The Cantor distribution
  *The generalized logistic distribution family
  *The Pearson distribution family
  *The phase-type distribution



#Works Cited
Joyce, D. (2016). Common probability distributionsI. Retrieved January 31, 2017, from http://aleph0.clarku.edu/~djoyce/ma218/distributions.pdf
Minitab. (n.d.). What is a continuous distribution? Retrieved January 31, 2017, from http://support.minitab.com/en-us/minitab/17/topic-library/basic-statistics-and-graphs/probability-distributions-and-random-data/basics/continuous-distribution/
Sonnier, R. (n.d.). Kolmogorov Smirnov VBA Code. Retrieved January 31, 2017, from http://rolfsonnier.typepad.com/blog/2012/07/download-kolmogorov-smirnov-vba-code.html
Weisstein, E. (2017, January 31). Discrete Distribution. Retrieved January 31, 2017, from http://mathworld.wolfram.com/DiscreteDistribution.html
Weisstein, E. (2017, January 31). Wolfram Mathworld. Retrieved January 31, 2017, from http://mathworld.wolfram.com/BernoulliDistribution.html
Wikipedia. (n.d.). Dirac comb. Retrieved January 31, 2017, from https://en.wikipedia.org/wiki/Dirac_comb
Wikipedia. (n.d.). Irwin–Hall distribution. Retrieved January 31, 2017, from https://en.wikipedia.org/wiki/Irwin%E2%80%93Hall_distribution
Wikipedia. (n.d.). Kent distribution. Retrieved January 31, 2017, from https://en.wikipedia.org/wiki/Kent_distribution
Wikipedia. (n.d.). Kumaraswamy distribution. Retrieved January 31, 2017, from https://en.wikipedia.org/wiki/Kumaraswamy_distribution
Wikipedia. (n.d.). List of Probability Distributions. Retrieved January 31, 2017, from https://en.wikipedia.org/wiki/List_of_probability_distributions
Wikipedia. (n.d.). Logit-normal distribution. Retrieved January 31, 2017, from https://en.wikipedia.org/wiki/Logit-normal_distribution
Wikipedia. (n.d.). Raised cosine distribution. Retrieved January 31, 2017, from https://en.wikipedia.org/wiki/Raised_cosine_distribution
Wikipedia. (n.d.). Triangular distribution. Retrieved January 31, 2017, from https://en.wikipedia.org/wiki/Triangular_distribution
Wikipedia. (n.d.). Uniform distribution (continuous). Retrieved January 31, 2017, from https://en.wikipedia.org/wiki/Uniform_distribution_(continuous)
Wikipedia. (n.d.). von Mises distribution. Retrieved January 31, 2017, from https://en.wikipedia.org/wiki/Von_Mises_distribution
Wikipedia. (n.d.). Wrapped asymmetric Laplace distribution. Retrieved January 31, 2017, from https://en.wikipedia.org/wiki/Wrapped_asymmetric_Laplace_distribution
Wikipedia. (n.d.). Wrapped Cauchy distribution. Retrieved January 31, 2017, from https://en.wikipedia.org/wiki/Wrapped_Cauchy_distribution
Wikipedia. (n.d.). Wrapped exponential distribution. Retrieved January 31, 2017, from https://en.wikipedia.org/wiki/Wrapped_exponential_distribution
Wikipedia. (n.d.). Wrapped normal distribution. Retrieved January 31, 2017, from https://en.wikipedia.org/wiki/Wrapped_normal_distribution
Zaiontz, C. (n.d.). Shapiro-Wilk Expanded Test. Retrieved January 31, 2017, from http://www.real-statistics.com/tests-normality-and-symmetry/statistical-tests-normality-symmetry/shapiro-wilk-expanded-test/
Zaiontz, C. (n.d.). Shapiro-Wilk Original Test. Retrieved January 31, 2017, from http://www.real-statistics.com/tests-normality-and-symmetry/statistical-tests-normality-symmetry/shapiro-wilk-test/


		



0Document102/13/2017



Document10Page 14






